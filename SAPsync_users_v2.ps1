#Скрипт управление пользователем в АД:
#
#в качестве параметра принимает JSON объект по этому пользователю
#
#ищет переданного пользователя в АД сначала по табельному номеру,
#затем по ФИО
#
#Синхронизирует поля:
# - ФИО
# - Табельный номер
# - Должность
# - Подразделение
# - Организация
# 
# v2.0 + Поддержка нескольких организаций
#        Убрана обертка вокруг БД Инвентаризации для взаимодействия с этими скриптами
#        Теперь все взаимодействие ведется через каноничный REST API            
# v1.5 + Городской номер приводится к тому же формату, что и мобильный
# v1.4 + Синхронизация почты, внутреннего номера телефона, мобильного номера телефона 
#        в сторону САП 
#      + мобильные телефонные номера сначала приводятся к формату +7(ХХХ)ХХХ-ХХХХ
# v1.3 Хранение табельного номера переведено на EmployeeID
# v1.2 Улучшена корректировака имен. В т.ч. переименование самого объекта
# v1.1 Улучшена корректировка имени. 
#      Пользователя можно искать указав в качестве имени табельный номер
# v1.0 Initial commit


. "$($PSScriptRoot)\..\config.priv.ps1"


#выводит сообщение в лог файл
function Log()
{
	param
	(
		[string]$msg
	)

    "$msg"
	$now = Get-Date
	"$(Get-Date) $msg" | Out-File -filePath $sync_logfile -append -encoding Default
}


function correctMobile() {
	param (
		[string]$number
	)
	$original=$number

	#убираем пробелы
	$number=$number.Replace(' ','').Replace('-','').Replace('.','')
	#Log($original+": clean ["+$number+"]")

	#проверяем что цифр 11	
	if ( -not ($number.Replace('+','').Replace('(','').Replace(')','').Length -eq 11)) {
		#Log($original+": numbers ["+$number.Replace('+','').Replace('(','').Replace(')','')+"]")
		#Log($original+": numberscount ["+$number.Replace('+','').Replace('(','').Replace(')','').Length+"]")
		return $original
	}
	
	#8XXX -> 7XXX
	if ($number.Substring(0,1) -eq "8") {
		$number="7"+$number.Substring(1)
	}

	#7XXX -> +7XXX
	if ($number.Substring(0,1) -eq "7") {
		$number="+"+$number
	}
	#Log($original+": country correct ["+$number+"]")

	#проверяем что скобочки есть и они расставлены правильно
	$leftBracket=$number.IndexOf("(")
	$rightBracket=$number.IndexOf(")")
	if ( ($leftBracket -lt 0) -or ($rightBracket -lt 0) -or ($rightBracket -lt $leftBracket) ) {
		$number=$number.Replace('(','').Replace(')','')
		$countryCode=$number.Substring(0,2);
		$cityCode=$number.Substring(2,3);
		$localCode=$number.Substring(5);
		$number=$countryCode+'('+$cityCode+')'+$localCode
		$rightBracket=$number.IndexOf(")")
	}

	#Log($original+": brackets correct ["+$number+"]")

	#проверяем знак тире
	$minusLeft=$number.Substring(0,$rightBracket+4)
	$minusRight=$number.Substring($rightBracket+4)
	
	return $minusLeft+'-'+$minusRight

}

#запись данных о пользователе в БД
function pushUserData() {
	param
	(
		[string]$id,
		[string]$field,
		[string]$value
	)
	$value=[System.Web.HttpUtility]::UrlEncode($value)

    $params = @{$field=$value;}

	try { 
        #Invoke-WebRequest -Uri "$($inventory_RESTapi_URL)/users/$($id)" -Method POST -Body $params
        Log("$($inventory_RESTapi_URL)/users/$($id)");
		Log("Success")
	} catch {
		Log("Error")
	}
}

#загрузить пользователя из Инвентаризации через REST API
function FindUser() {
	param
	(
		[object]$user
	)
    #Мы можем реализовать три сценария:
    #Если у нас есть организация и табельный - ищем конкретного
    #Если у нас есть только табельный - считаем что организация=1 и гото 1
    #Если нету ничего - ищем по ФИО
    #следовательно главно это есть табельный или нет:
    #Log("$($user.employeeNumber):$($user.employeeNumber.Length)")
    if ($user.employeeNumber.Length -gt 0) {
        #Log("IDENTIFIED")
        #табельный есть:
        $org_id=$user.employeeID
        if ($org_id.Length -eq 0) {
            #если организация не заявлена, то первая
            #эта ситуация скорее всего возникнет при переходе от инвентаризации версии под одну организацию
            #к инвентаризации версии под множество. Когда БД уже с учетом организаций а АД еще нет
            $org_id=1
        }
        #запрос пользователя будет по организации и табельному
        $reqParams="num=$($user.employeeNumber)&org=$($org_id)"
    } else {
        $reqParams="name=$($user.displayName)"
    }
    $webReq="$($inventory_RESTapi_URL)/users/view?$($reqParams)&expand=ln,mn,fn"
    #Log($webReq)
	try { 
		$sap = ((invoke-WebRequest $webReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
	} catch {
		$err=$_.Exception.Response.StatusCode.Value__
        Log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] not found in SAP by $($webReq)")
        $sap='error'
    }
    return $sap
}

#обработка пользователя
function ParseUser() {
	param
	(
		[object]$user
	)
	#Выставляем флажки
	#Обновлять пользоватея в АД не надо
	$needUpdate = $false
	#Переименовывать пользоватея в АД не надо
	$needRename = $false

	
	$sap = findUser($user)
	#Если пользователь нашелся
	if ( -not ($sap -eq "error")) {
		if ($sap.Uvolen -eq "1") {
			#Уволенных увольняем
			Log($user.sAMAccountname+ ": user dissmissed! Deactivation needed!")
			#c:\tools\usermanagement\usr_dismiss.cmd $user.sAMAccountname
		} else {
			#Грузим Ф И О по оттдельности
			$fn=($sap.fn).trim()
			$mn=($sap.mn).trim()
			$ln=($sap.ln).trim()
			$gn=($fn+" "+$mn).trim()

			#проверка пользователя на совпадение "названия" с ФИО
			if (
				($user.name -ne $sap.Ename) -or
				($user.cn -ne $sap.Ename)
			){
				Log($user.sAMAccountname+": got Name ["+$user.displayName+"] instead of ["+$sap.Ename+"] - Object rename needed")
				$needRename = $true
			}


			#проверка Выводимого имени пользователя на совпадение с ФИО
			if (
				($user.displayName -ne $sap.Ename) 
			){
				Log($user.sAMAccountname+": got displayName ["+$user.displayName+"] instead of ["+$sap.Ename+"]")
				$user.displayName=$sap.Ename
				$needUpdate = $true
			}

			#проверка Имени и Фамилии пользователя на совпадение с Именем и Фамилией
			if (
				($user.givenName -ne $gn) -or
				($user.sn -ne $ln)
			){
				Log($user.sAMAccountname+": got firstName+lastName ["+$user.givenName+" "+$user.sn+"] instead of ["+$gn+" "+$ln+"]")
				$user.sn=$ln
				$user.givenName=$gn
				$needUpdate = $true
			}

            #Подразделение

			if (
				($sap.orgStruct.name.Length -gt 0) -and
				($user.department -ne $sap.orgStruct.name )
			){
				Log($user.sAMAccountname+": got Department ["+$user.department+"] instead of ["+$sap.orgStruct.name+"]")
				$user.department=$sap.orgStruct.name
				$needUpdate = $true
			}

            #Организация
			if (
				($sap.org.name.Length -gt 0 ) -and
				($user.company -ne $sap.org.name )
			){
				Log($user.sAMAccountname+": got Org ["+$user.company+"] instead of ["+$sap.org.name+"]")
				$user.company=$sap.org.name
				$needUpdate = $true
			}

            #Должность
			if (
				($sap.Doljnost.Length -gt 0) -and
				($user.title -ne $sap.Doljnost)
			){
				Log($user.sAMAccountname+": got Title ["+$user.title+"] instead of ["+$sap.Doljnost+"]")
				$user.title=$sap.Doljnost
				$needUpdate = $true
			}

            #табельный номмер
			if ($user.EmployeeNumber -ne $sap.employee_id ){
				Log($user.sAMAccountname+": got Numbr ["+$user.EmployeeNumber+"] instead of ["+$sap.employee_id+"]")
				$user.EmployeeNumber=$sap.employee_id
				$needUpdate = $true
			}
			
			if ($user.EmployeeID -ne $sap.org_id ){
				Log($user.sAMAccountname+": got orgID ["+$user.EmployeeID+"] instead of ["+$sap.org_id+"]")
				$user.EmployeeID=$sap.org_id
				$needUpdate = $true
			}

			#$correctedMobile=correctMobile($sap.data.Mobile)
			#отключаем корректировку номера, т.к. во первых в ямале несколько номеров вбито
			#во вторых они частично приведены к одному виду уже в БД
			#в третьих там встречаются МН номера
			$correctedMobile=$sap.Mobile
			if ([string]$user.mobile -ne [string]$correctedMobile) {
				#для поля мобильного делаем обработку на случай если оно стало пустым, т.к. это реальная ситуация
				if ($correctedMobile -eq "") {
					Log($user.sAMAccountname+": got incorrect mobile ["+$user.mobile+"] instead of [empty]")
					$tmpUser = Get-ADUser $user.DistinguishedName
					#Set-AdUser $tmpUser -Clear mobile
				} else {
					Log($user.sAMAccountname+": got incorrect mobile ["+$user.mobile+"] instead of ["+$correctedMobile+"]")
					$user.mobile=$correctedMobile
					$needUpdate = $true
				}
			}

			$correctedPhone=correctMobile($user.telephoneNumber)
			if ([string]$user.telephoneNumber -ne [string]$correctedPhone) {
				Log($user.sAMAccountname+": got incorrect telephoneNumber format ["+$user.telephoneNumber+"] instead of ["+$correctedPhone+"]")
				$user.telephoneNumber=$correctedPhone
				$needUpdate = $true
			}

			#http://inventory.yamalgazprom.local/web/api/phones/search-by-user?id=%D0%90%D0%9E%D0%97%D0%9F-00960
			#Запрашиваем номер телефона, привязанный к пользователю в Инвентаризации
			$webReqPh="$($inventory_RESTapi_URL)/phones/search-by-user?id=$($sap.id)"
			#$webReqPh
			try { 
				$sapPh = ((invoke-WebRequest $webReqPh -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
				#$sapPh
				if (
					($sapPh.length -gt 2 ) -and 
					($sapPh -ne $user.Pager)
				) {
					Log($user.sAMAccountname+": got Phone ["+$user.pager+"] instead of ["+$sapPh+"]")
					$user.pager=$sapPh
					$needUpdate = $true
				}
			} catch {
				$err=$_.Exception.Response.StatusCode.Value__
			}


			if ($needUpdate) {
				$user 
				#Set-AdUser -Instance $user 
				Log($user.sAMAccountname+": changes pushed to AD")
				#exit(0)
			}
			if ($needRename) {
				Log($user.sAMAccountname+": AdObject renaming to "+$sap.Ename)
				#Rename-AdObject -Identity $user -Newname $sap.Ename
				Log($user.sAMAccountname+": AdObject renamed to "+$sap.Ename)
			}
			$webWriteReq="$($inventory_write_api_URL)&IPernr=$($sap.data.Pernr)"
			
			#push данных обратно в БД
			if ([string]$user.mail -ne [string]$sap.Email) {
				Log($user.sAMAccountname+": got SAP email ["+$sap.Email+"] instead of ["+$user.mail+"]")
				pushUserData $sap.id Email $user.mail
			}
			if ([string]$user.pager -ne [string]$sap.Phone) {
				Log($user.sAMAccountname+": got SAP phone ["+$sap.Phone+"] instead of ["+$user.pager+"]")
				pushUserData $sap.id Phone $user.pager
			}
			if ([string]$user.sAMAccountname -ne [string]$sap.Login) {
				Log($user.sAMAccountname+": got SAP Login ["+$sap.Login+"] instead of ["+$user.sAMAccountname+"]")
				pushUserData $sap.id Login $user.sAMAccountname
			}
		}
	}
}

Import-Module ActiveDirectory

$users = Get-ADUser -Filter {enabled -eq $true} -SearchBase $u_OUDN -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber
$u_count = $users | measure 
Write-Host "Users to sync: " $u_count.Count

foreach($user in $users) {
    #$user
	ParseUser ($user)
    #exit
}
