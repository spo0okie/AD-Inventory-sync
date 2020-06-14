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

# v2.2 + Учет даты увольния при увольнении
# v2.1 + Обработка нескольких телефонов через запятую, ограничение длины полей
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

#как посмотреть лимит на длину поля? например для mobile вот так:
#dsquery * "cn=Schema,cn=Configuration,dc=yamalgazprom,dc=local" -Filter "(LDAPDisplayName=mobile)" -attr rangeUpper

#у нас были затыки с полями
#mobile (64)
#title (128)

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\lib_funcs.ps1"


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


#запись данных о пользователе в БД
function pushUserData() {
	param
	(
		[string]$id,
		[string]$field,
		[string]$value
	)
	#$value=[System.Web.HttpUtility]::UrlEncode($value)

	$params = @{$field=$value;}

    if ($write_inventory) {
        $uri="$($inventory_RESTapi_URL)/users/$($id)"
	    try { 
    
            $result=Invoke-WebRequest -Uri $uri -Method PUT -Body $params -UseBasicParsing
            #Log("$($inventory_RESTapi_URL)/users/$($id)")
		    Log("Success")
	    } catch {
		    Log("Error: $($_.Exception.Response.StatusCode.Value__): $($_.Exception.Message)")
            $_.Exception.Response
            $_.Exception.content

	    }
    } else {
        #log ("invPush: skip user #$id $field = $value INV: RO mode")
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

    $org_id=$user.employeeID
    if (($org_id.Length -eq 0) -or ( -not $mutiorg_support)) {
        #если организация не заявлена, то первая
        #эта ситуация скорее всего возникнет при переходе от инвентаризации версии под одну организацию
        #к инвентаризации версии под множество. Когда БД уже с учетом организаций а АД еще нет
        $org_id=1
    }

    #Log("$($user.employeeNumber):$($user.employeeNumber.Length)")
    if ($user.employeeNumber.Length -gt 0) {
        #Log("IDENTIFIED")
        #табельный есть:
        #запрос пользователя будет по организации и табельному
        $reqParams="num=$($user.employeeNumber)&org=$($org_id)"
    } elseif ($user.displayName.Length -gt 0) {
        $reqParams="name=$($user.displayName)"
    } elseif ($user.login.sAMAccountname -gt 0) {
        $reqParams="login=$($user.sAMAccountname)"
    } else {
        log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] - don't know how to search 0_o")
        return 'error'
    }


    $webReq="$($inventory_RESTapi_URL)/users/view?$($reqParams)&expand=ln,mn,fn,orgStruct"
    #Log($webReq)

    #пробуем найти нашего сотрудника	
    try { 
        $sap = ((invoke-WebRequest $webReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
    } catch {
        #неудача!
        $err=$_.Exception.Response.StatusCode.Value__
        Log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] not found in SAP by $($webReq)")
        #Действуем по плану Б:
        #В имени сотрудника может быть вбит табельный. Ищем его по табельному из имени:
        $reqParams="num=$($user.displayName)&org=$($org_id)"
        $webReq="$($inventory_RESTapi_URL)/users/view?$($reqParams)&expand=ln,mn,fn"
        try { 
            $sap = ((invoke-WebRequest $webReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
        } catch {
            $err=$_.Exception.Response.StatusCode.Value__
            Log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] not found in SAP by $($webReq)")
            $sap='error'
        }
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
	#Переименовывать пользователя в АД не надо
	$needRename = $false
	#Увольнять пользователя в АД не надо
	$needDismiss = $false

	
	$sap = FindUser($user)
	#Если пользователь нашелся
	if ( -not ($sap -eq "error")) {
		
		if ($sap.Uvolen -eq "1") {
			#смотрим когда уволен
			if ($sap.resign_date.Length -gt 0) {
				$resign_date=[datetime]::parseexact($sap.resign_date, $dateformat_SAP, $null)
				#уже уволен?
				if ((Get-Date) -gt $resign_date) {
					$needDismiss = $true
				}
			}
		}
		if ($needDismiss) {
			#Уволенных увольняем
			#проверяем исключения уволенных
			if ($auto_dismiss_exclude -eq $user.sAMAccountname) {
				Log($user.sAMAccountname+ ": user dissmissed! Deactivation disabled (exclusion list)!")
			} else {
				if ($auto_dismiss) {
					Log($user.sAMAccountname+ ": user dissmissed! Deactivating")
					#c:\tools\usermanagement\usr_dismiss.cmd $user.sAMAccountname
				} else {
					Log($user.sAMAccountname+ ": user dissmissed! Deactivation needed!")
				}
			}
			
		} else {
			#Log($user.sAMAccountname+ ": user active")
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
			$title=$sap.Doljnost
			if ($title.Length -gt 128) {
				#ограничение длины поля
				$title=$title.Substring(0,128)
			}
			if (
				($title -gt 0) -and
				($user.title -ne $title)
			){
				Log($user.sAMAccountname+": got Title ["+$user.title+"] instead of ["+$title+"]")
				$user.title=$title
				$needUpdate = $true
			}

            #табельный номер
			if ($user.EmployeeNumber -ne $sap.employee_id ){
				Log($user.sAMAccountname+": got Numbr ["+$user.EmployeeNumber+"] instead of ["+$sap.employee_id+"]")
				$user.EmployeeNumber=$sap.employee_id
				$needUpdate = $true
			}
			
            #ID организации
            if ($mutiorg_support) {
			    if ($user.EmployeeID -ne $sap.org_id ){
				    Log($user.sAMAccountname+": got orgID ["+$user.EmployeeID+"] instead of ["+$sap.org_id+"]")
				    $user.EmployeeID=$sap.org_id
				    $needUpdate = $true
			    }
            }


            if ($mobile_from_SAP) {
			    #$correctedMobile=correctMobile($sap.Mobile)
			    #отключаем корректировку номера, т.к. во первых в ямале несколько номеров вбито
			    #во вторых они частично приведены к одному виду уже в БД
			    #в третьих там встречаются МН номера
			    $correctedMobile= correctPhonesList($sap.Mobile)
			    if ([string]$user.mobile -ne [string]$correctedMobile) {
				    #для поля мобильного делаем обработку на случай если оно стало пустым, т.к. это реальная ситуация
				    if ($correctedMobile -eq "") {
					    Log($user.sAMAccountname+": got incorrect mobile ["+$user.mobile+"] instead of [empty]")
					    $tmpUser = Get-ADUser $user.DistinguishedName
                        if ($write_AD) {
    					    Set-AdUser $tmpUser -Clear mobile
                        }
				    } else {
					    Log($user.sAMAccountname+": got incorrect mobile ["+$user.mobile+"] instead of ["+$correctedMobile+"]")
					    $user.mobile=$correctedMobile
					    $needUpdate = $true
				    }
			    }
            }

			$correctedPhone=correctMobile($user.telephoneNumber)
			if ([string]$user.telephoneNumber -ne [string]$correctedPhone) {
				Log($user.sAMAccountname+": got incorrect telephoneNumber format ["+$user.telephoneNumber+"] instead of ["+$correctedPhone+"]")
				$user.telephoneNumber=$correctedPhone
				$needUpdate = $true
			}

			if ($phone_from_SAP ){
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
            }


			if ($needUpdate) {
                if ($write_AD) {
    				$user 
				    Set-AdUser -Instance $user 
	    			Log($user.sAMAccountname+": changes pushed to AD")
                } else {
	    			Log($user.sAMAccountname+": AD push skipped: AD RO mode")
                }
				#exit(0)
			}
			if ($needRename) {
                if ($write_AD) {
	    			Log($user.sAMAccountname+": AdObject renaming to "+$sap.Ename)
    				Rename-AdObject -Identity $user -Newname $sap.Ename
		    		Log($user.sAMAccountname+": AdObject renamed to "+$sap.Ename)
                } else {
	    			Log($user.sAMAccountname+": rename $($user.sAMAccountname) -> $($sap.Ename) skipped: AD RO mode")
                }
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
            if ( -not $mobile_from_SAP) {
			    $correctedMobile= correctMobile($user.mobile)
			    if ([string]$sap.Mobile -ne [string]$correctedMobile) {
					Log($user.sAMAccountname+": got SAP mobile ["+$sap.mobile+"] instead of ["+$correctedMobile+"]")
					pushUserData $sap.id Mobile $correctedMobile
			    }
            }
		}
	} else {
		# Log($user.sAMAccountname+": Skip: got SAP error")
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
