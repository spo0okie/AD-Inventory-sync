#������ ���������� ������������� � ��:
#
#� �������� ��������� ��������� JSON ������ �� ����� ������������
#
#���� ����������� ������������ � �� ������� �� ���������� ������,
#����� �� ���
#
#�������������� ����:
# - ���
# - ��������� �����
# - ���������
# - �������������
# - �����������
# 
# v2.0 + ��������� ���������� �����������
#        ������ ������� ������ �� �������������� ��� �������������� � ����� ���������
#        ������ ��� �������������� ������� ����� ���������� REST API            
# v1.5 + ��������� ����� ���������� � ���� �� �������, ��� � ���������
# v1.4 + ������������� �����, ����������� ������ ��������, ���������� ������ �������� 
#        � ������� ��� 
#      + ��������� ���������� ������ ������� ���������� � ������� +7(���)���-����
# v1.3 �������� ���������� ������ ���������� �� EmployeeID
# v1.2 �������� �������������� ����. � �.�. �������������� ������ �������
# v1.1 �������� ������������� �����. 
#      ������������ ����� ������ ������ � �������� ����� ��������� �����
# v1.0 Initial commit


. "$($PSScriptRoot)\..\config.priv.ps1"


#������� ��������� � ��� ����
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

	#������� �������
	$number=$number.Replace(' ','').Replace('-','').Replace('.','')
	#Log($original+": clean ["+$number+"]")

	#��������� ��� ���� 11	
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

	#��������� ��� �������� ���� � ��� ����������� ���������
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

	#��������� ���� ����
	$minusLeft=$number.Substring(0,$rightBracket+4)
	$minusRight=$number.Substring($rightBracket+4)
	
	return $minusLeft+'-'+$minusRight

}

#������ ������ � ������������ � ��
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

#��������� ������������ �� �������������� ����� REST API
function FindUser() {
	param
	(
		[object]$user
	)
    #�� ����� ����������� ��� ��������:
    #���� � ��� ���� ����������� � ��������� - ���� �����������
    #���� � ��� ���� ������ ��������� - ������� ��� �����������=1 � ���� 1
    #���� ���� ������ - ���� �� ���
    #������������� ������ ��� ���� ��������� ��� ���:
    #Log("$($user.employeeNumber):$($user.employeeNumber.Length)")
    if ($user.employeeNumber.Length -gt 0) {
        #Log("IDENTIFIED")
        #��������� ����:
        $org_id=$user.employeeID
        if ($org_id.Length -eq 0) {
            #���� ����������� �� ��������, �� ������
            #��� �������� ������ ����� ��������� ��� �������� �� �������������� ������ ��� ���� �����������
            #� �������������� ������ ��� ���������. ����� �� ��� � ������ ����������� � �� ��� ���
            $org_id=1
        }
        #������ ������������ ����� �� ����������� � ����������
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

#��������� ������������
function ParseUser() {
	param
	(
		[object]$user
	)
	#���������� ������
	#��������� ����������� � �� �� ����
	$needUpdate = $false
	#��������������� ����������� � �� �� ����
	$needRename = $false

	
	$sap = findUser($user)
	#���� ������������ �������
	if ( -not ($sap -eq "error")) {
		if ($sap.Uvolen -eq "1") {
			#��������� ���������
			Log($user.sAMAccountname+ ": user dissmissed! Deactivation needed!")
			#c:\tools\usermanagement\usr_dismiss.cmd $user.sAMAccountname
		} else {
			#������ � � � �� ������������
			$fn=($sap.fn).trim()
			$mn=($sap.mn).trim()
			$ln=($sap.ln).trim()
			$gn=($fn+" "+$mn).trim()

			#�������� ������������ �� ���������� "��������" � ���
			if (
				($user.name -ne $sap.Ename) -or
				($user.cn -ne $sap.Ename)
			){
				Log($user.sAMAccountname+": got Name ["+$user.displayName+"] instead of ["+$sap.Ename+"] - Object rename needed")
				$needRename = $true
			}


			#�������� ���������� ����� ������������ �� ���������� � ���
			if (
				($user.displayName -ne $sap.Ename) 
			){
				Log($user.sAMAccountname+": got displayName ["+$user.displayName+"] instead of ["+$sap.Ename+"]")
				$user.displayName=$sap.Ename
				$needUpdate = $true
			}

			#�������� ����� � ������� ������������ �� ���������� � ������ � ��������
			if (
				($user.givenName -ne $gn) -or
				($user.sn -ne $ln)
			){
				Log($user.sAMAccountname+": got firstName+lastName ["+$user.givenName+" "+$user.sn+"] instead of ["+$gn+" "+$ln+"]")
				$user.sn=$ln
				$user.givenName=$gn
				$needUpdate = $true
			}

            #�������������

			if (
				($sap.orgStruct.name.Length -gt 0) -and
				($user.department -ne $sap.orgStruct.name )
			){
				Log($user.sAMAccountname+": got Department ["+$user.department+"] instead of ["+$sap.orgStruct.name+"]")
				$user.department=$sap.orgStruct.name
				$needUpdate = $true
			}

            #�����������
			if (
				($sap.org.name.Length -gt 0 ) -and
				($user.company -ne $sap.org.name )
			){
				Log($user.sAMAccountname+": got Org ["+$user.company+"] instead of ["+$sap.org.name+"]")
				$user.company=$sap.org.name
				$needUpdate = $true
			}

            #���������
			if (
				($sap.Doljnost.Length -gt 0) -and
				($user.title -ne $sap.Doljnost)
			){
				Log($user.sAMAccountname+": got Title ["+$user.title+"] instead of ["+$sap.Doljnost+"]")
				$user.title=$sap.Doljnost
				$needUpdate = $true
			}

            #��������� ������
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
			#��������� ������������� ������, �.�. �� ������ � ����� ��������� ������� �����
			#�� ������ ��� �������� ��������� � ������ ���� ��� � ��
			#� ������� ��� ����������� �� ������
			$correctedMobile=$sap.Mobile
			if ([string]$user.mobile -ne [string]$correctedMobile) {
				#��� ���� ���������� ������ ��������� �� ������ ���� ��� ����� ������, �.�. ��� �������� ��������
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
			#����������� ����� ��������, ����������� � ������������ � ��������������
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
			
			#push ������ ������� � ��
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
