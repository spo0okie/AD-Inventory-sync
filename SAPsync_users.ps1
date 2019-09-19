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

	$now = Get-Date
	"$(Get-Date) $msg" | Out-File -filePath $logfile -append -encoding Default
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

#запись данных о пользователе в САП
function pushUserData() {
	param
	(
		[string]$pernr,
		[string]$type,
		[string]$value
	)
	$value=[System.Web.HttpUtility]::UrlEncode($value)
	$webWriteReq="$($inventory_write_api_URL)&IPernr=$($pernr)&ISubty=$($type)&IValue=$($value)"
	Log("Requesting " +$webWriteReq)
	$sapResult = ((invoke-WebRequest $webWriteReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
	if ($sapResult.OOk -eq "X") {
		Log("Success")
	} else {
		Log("Error")
	}
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

	#Запрашиваем объект пользователя из HR БД
	$webReq="$($inventory_api_URL)?name=$($user.displayName)&number=$($user.EmployeeNumber)"
	$sap = ((invoke-WebRequest $webReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
	#Если пользователь нашелся
	if ($sap.result -eq "OK") {
		if ($sap.data.Uvolen -eq "1") {
			#Уволенных увольняем
			Log($user.sAMAccountname+ ": user dissmissed! Deactivation needed!")
			c:\tools\usermanagement\usr_dismiss.cmd $user.sAMAccountname
		} else {
			#Грузим Ф И О по оттдельности
			$fn=($sap.data.Vorna).trim()
			$mn=($sap.data.Midnm).trim()
			$ln=($sap.data.Nachn).trim()
			$gn=($fn+" "+$mn).trim()

			#проверка пользователя на совпадение "названия" с ФИО
			if (
				($user.name -ne $sap.data.Ename) -or
				($user.cn -ne $sap.data.Ename)
			){
				Log($user.sAMAccountname+": got Name ["+$user.displayName+"] instead of ["+$sap.data.Ename+"] - Object rename needed")
				$needRename = $true
			}


			#проверка Выводимого имени пользователя на совпадение с ФИО
			if (
				($user.displayName -ne $sap.data.Ename) 
			){
				Log($user.sAMAccountname+": got displayName ["+$user.displayName+"] instead of ["+$sap.data.Ename+"]")
				$user.displayName=$sap.data.Ename
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

			if (
				($sap.data.Orgtx.Length > 0) -and
				($user.department -ne $sap.data.Orgtx )
			){
				Log($user.sAMAccountname+": got Department ["+$user.department+"] instead of ["+$sap.data.Orgtx+"]")
				$user.department=$sap.data.Orgtx
				$needUpdate = $true
			}

			if (
				($sap.data.Organization.Length > 0 ) -and
				($user.company -ne $sap.data.Organization )
			){
				Log($user.sAMAccountname+": got Org ["+$user.company+"] instead of ["+$sap.data.Organization+"]")
				$user.company=$sap.data.Organization
				$needUpdate = $true
			}

			if (
				($sap.data.Doljnost.Length > 0) -and
				($user.title -ne $sap.data.Doljnost)
			){
				Log($user.sAMAccountname+": got Title ["+$user.title+"] instead of ["+$sap.data.Doljnost+"]")
				$user.title=$sap.data.Doljnost
				$needUpdate = $true
			}

			if ($user.EmployeeNumber -ne $sap.data.Pernr ){
				Log($user.sAMAccountname+": got Numbr ["+$user.EmployeeNumber+"] instead of ["+$sap.data.Pernr+"]")
				$user.EmployeeNumber=$sap.data.Pernr
				$needUpdate = $true
			}
			
			#$correctedMobile=correctMobile($sap.data.Mobile)
			#отключаем корректировку номера, т.к. во первых в ямале несколько номеров вбито
			#во вторых они частично приведены к одному виду уже в БД
			#в третьих там встречаются МН номера
			$correctedMobile=$sap.data.Mobile
			if ([string]$user.mobile -ne [string]$correctedMobile) {
				#для поля мобильного делаем обработку на случай если оно стало пустым, т.к. это реальная ситуация
				if ($correctedMobile -eq "") {
					Log($user.sAMAccountname+": got incorrect mobile ["+$user.mobile+"] instead of [empty]")
					$tmpUser = Get-ADUser $user.DistinguishedName
					Set-AdUser $tmpUser -Clear mobile
					#Set-AdUser -Instance $tmpUser -Clear mobile
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

			if ($needUpdate) {
				$user 
				Set-AdUser -Instance $user 
				Log($user.sAMAccountname+": changes pushed to AD")
				#exit(0)
			}
			if ($needRename) {
				Log($user.sAMAccountname+": AdObject renaming to "+$sap.data.Ename)
				Rename-AdObject -Identity $user -Newname $sap.data.Ename
				Log($user.sAMAccountname+": AdObject renamed to "+$sap.data.Ename)
			}
			$webWriteReq="$($inventory_write_api_URL)&IPernr=$($sap.data.Pernr)"
			
			#push данных обратно в БД
			if ([string]$user.mail -ne [string]$sap.data.Email) {
				Log($user.sAMAccountname+": got SAP email ["+$sap.data.Email+"] instead of ["+$user.mail+"]")
				pushUserData $sap.data.Pernr Email $user.mail
			}
			if ([string]$user.pager -ne [string]$sap.data.Phone) {
				Log($user.sAMAccountname+": got SAP phone ["+$sap.data.Phone+"] instead of ["+$user.pager+"]")
				pushUserData $sap.data.Pernr Phone $user.pager
			}
			if ([string]$user.sAMAccountname -ne [string]$sap.data.Login) {
				Log($user.sAMAccountname+": got SAP Login ["+$sap.data.Login+"] instead of ["+$user.sAMAccountname+"]")
				pushUserData $sap.data.Pernr Login $user.sAMAccountname
			}
		}
	} elseif ($sap.code -eq "1" ) {
		Log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] not found in SAP")
	} elseif ($sap.code -eq "2" ) {
		Log("ERROR: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] got more than one employments")
	}
}

Import-Module ActiveDirectory

$users = Get-ADUser -Filter {enabled -eq $true} -SearchBase $u_OUDN -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,EmployeeNumber,mail,pager,mobile,telephoneNumber
$u_count = $users | measure 
Write-Host "Найдено пользователей по запросу: " $u_count.Count

foreach($user in $users) {
	ParseUser ($user)
}

