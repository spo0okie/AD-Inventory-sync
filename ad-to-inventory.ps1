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

#как посмотреть лимит на длину поля? например для mobile вот так:
#dsquery * "cn=Schema,cn=Configuration,dc=Domain,dc=local" -Filter "(LDAPDisplayName=mobile)" -attr rangeUpper

#у нас были затыки с полями
#mobile (64)
#title (128)

$global:DEBUG_MODE=1

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_funcs.ps1"
. "$($PSScriptRoot)\..\libs.ps1\lib_inventory.ps1"



#загрузить пользователя из Инвентаризации через REST API Для синхронизации Инвентори -> АД
function FindUser() {
	param
	(
		[object]$user
	)
    #Ищем пользователя последовательно:

    #ищем по UID
    $invUser=getInventoryObj 'users' '' @{
        uid=$user.objectGUID;
        expand='orgStruct';
    }
    if ($invUser -is [PSCustomObject]) {return $invUser}

    #по Login
    $invUser=getInventoryObj 'users' '' @{
        login=$user.sAMAccountname;
        expand='orgStruct';
    }
    if ($invUser -is [PSCustomObject]) {return $invUser}

    #по имени
    $invUser=getInventoryObj 'users' '' @{
        name=$user.displayName;
        expand='orgStruct';
    }
    if ($invUser -is [PSCustomObject]) {return $invUser}

    return $false
}

#получает организацию по имени
function fetchOrganization() {
	param (
		[string]$orgName
	)
	$org=(getInventoryObj 'partners' $orgName)
	return $org
}

#подразделение по имени и организации
function fetchDepartment() {
	param (
		[int]$org_id,
		[string]$departmentName
	)
	$department=getInventoryObj 'org-struct' $departmentName @{org_id=$org_id}
	if ( -not $department ) {
		spooLog "Department $($departmentName) not found"
		if ($global:write_inventory) {
	        $dep=@{ 
				org_id = $org_id;
			    name = $departmentName;
			}
			$department= setInventoryData 'org-struct' $dep		
		} else {
			spooLog ("inventory RO: Skip department $($departmentName) creation")
			return $false
		}
	}
    return $department
}

#обработка пользователя
function ParseUser() {
	param
	(
		[object]$user
	)

    #список изменений пустой
	$changes=@{}

	#ищем в инвентори
    $invUser = FindUser($user)
	#Если пользователь не нашелся
	if ( -not ( $invUser -is [PSCustomObject] )) {
		$invUser=@{
            Persg=1;
            id=-1;
            Uvolen=0;
            nosync=0;
        }
	}

    if ($invUser.nosync) {
        warningLog ($user.sAMAccountname+": synchronization disabled on Inventory side. SKIP");
        return
    }


	#uid
	if (([string]$user.objectGUID -ne [string]$invUser.uid)){
		spooLog($user.sAMAccountname+": got UID ["+$invUser.uid+"] instead of ["+$user.objectGUID+"]")
		$changes['uid']=$user.objectGUID
	}

    #Login
	if ([string]$user.sAMAccountname -ne [string]$invUser.Login) {
		spooLog($user.sAMAccountname+": got Inventory Login ["+$invUser.Login+"] instead of ["+$user.sAMAccountname+"]")
		$changes['Login']=$user.sAMAccountname
	}

    #e-mail (приводятся к строке, т.к. бывает NULL который не равен пустой строке)
	if ([string]$user.mail -ne [string]$invUser.Email) {
		spooLog($user.sAMAccountname+": got Inventory email ["+$invUser.Email+"] instead of ["+$user.mail+"]")
		$changes['Email']=$user.mail
	}
		
	#проверка пользователя на совпадение "названия" с ФИО
	if (
		([string]$user.displayName -ne [string]$invUser.Ename) 
	){
		spooLog($user.sAMAccountname+": got Name ["+$invUser.Ename+"] instead of ["+$user.displayName+"]")
		$changes['Ename']=$user.displayName;
	}

	#Должность
	if (
		([string]$user.title -ne [string]$invUser.Doljnost)
	){
		spooLog($user.sAMAccountname+": got Title ["+$invUser.Doljnost+"] instead of ["+$user.title+"]")
		$changes['Doljnost']=$user.title
	}

	#Мобильник
	if (
		([string]$user.mobile -ne [string]$invUser.mobile)
	){
		spooLog($user.sAMAccountname+": got Mobile ["+$invUser.mobile+"] instead of ["+$user.mobile+"]")
		$changes['Mobile']=$user.mobile
	}

	#Внутренний номер
	if (
		([string]$user.pager -ne [string]$invUser.phone)
	){
		spooLog($user.sAMAccountname+": got Phone ["+$invUser.phone+"] instead of ["+$user.pager+"]")
		$changes['Phone']=$user.pager
	}

	#Организация
	if (
		(([string]$user.company).Length -gt 0 ) #если в АД проставлена организация, то
	){
		$org = fetchOrganization($user.company); #ищем эту организацию в инвентори
		if ($org) { #если нашли, то работаем
			if ($invUser.org_id -ne $org.id) {
	    		spooLog($user.sAMAccountname+": got Org ["+$user.org.bname+"] instead of ["+$org.bname+"]")
				$changes['org_id']=$org.id
			}

			#Подразделение
			if (
				([string]$user.department -ne [string]$invUser.orgStruct.name )
			){
				spooLog($user.sAMAccountname+": got Department ["+$user.department+"] instead of ["+$invUser.orgStruct.name+"]")
                $dep = fetchDepartment $org.id $user.department
                if ( -not $dep) {
                    errorLog ("Failed fetch department $($user.department) for $($user.company)");
                } else {
    				$changes['Orgeh']=$dep.id
                }
			}

			
		} else {
			spooLog($user.company+": Org not found in inventory. Please create!");
		}
	}

    #признак увольнения
    if ($user.Enabled) {
        $uvolen=0         
    } else {
        $uvolen=1
    }

	if (
		($uvolen -ne $invUser.Uvolen)
	){
		spooLog($user.sAMAccountname+": got Fired status ["+$invUser.Uvolen+"] instead of ["+$uvolen+"]")
		$changes['Uvolen']=$uvolen
	}

    if ($changes.Count) {
        if ($invUser.id -ge 0) {
            spooLog ("updating user $($user.sAMAccountName)")
        } else {
            spooLog ("creating user $($user.sAMAccountName)")
        }
        $changes['Persg']=$invUser.Persg
        $changes['id']=$invUser.id
        $changes['Uvolen']=$uvolen
        $changes['nosync']=0;       #раз уж пишем, значит синхронизация не отключена
		$changes
		if ($global:write_inventory) {
	        setInventoryData 'users' $changes
		} else {
			spooLog ("inventory RO: Skip user $($user.sAMAccountName)")
		}
    }
			
}


Import-Module ActiveDirectory

#Пользователь передан конкретный
if ($args.Length -gt 0) {
	$user = Get-ADUser $args[0] -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber,adminDescription,enabled
	ParseUser ($user)
} else {
#Ничего не передано, обходим все OU
    foreach ($OUDN in $ad2inventory_OUDNs) {
		$users = Get-ADUser -Filter * -SearchBase $OUDN -properties Name,cn,sn,givenName,DisplayName,sAMAccountname,company,department,title,employeeNumber,employeeID,mail,pager,mobile,telephoneNumber,adminDescription,enabled
		$u_count = $users | measure 
		Write-Host "Users to sync in $($OUDN): " $u_count.Count

		foreach($user in $users) {
    		#$user
			ParseUser ($user)
    		#exit
		}
	}
}

