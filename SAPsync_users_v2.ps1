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

# v2.2 + ���� ���� �������� ��� ����������
# v2.1 + ��������� ���������� ��������� ����� �������, ����������� ����� �����
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

#��� ���������� ����� �� ����� ����? �������� ��� mobile ��� ���:
#dsquery * "cn=Schema,cn=Configuration,dc=yamalgazprom,dc=local" -Filter "(LDAPDisplayName=mobile)" -attr rangeUpper

#� ��� ���� ������ � ������
#mobile (64)
#title (128)

. "$($PSScriptRoot)\..\config.priv.ps1"
. "$($PSScriptRoot)\..\lib_funcs.ps1"


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


#������ ������ � ������������ � ��
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

    $org_id=$user.employeeID
    if (($org_id.Length -eq 0) -or ( -not $mutiorg_support)) {
        #���� ����������� �� ��������, �� ������
        #��� �������� ������ ����� ��������� ��� �������� �� �������������� ������ ��� ���� �����������
        #� �������������� ������ ��� ���������. ����� �� ��� � ������ ����������� � �� ��� ���
        $org_id=1
    }

    #Log("$($user.employeeNumber):$($user.employeeNumber.Length)")
    if ($user.employeeNumber.Length -gt 0) {
        #Log("IDENTIFIED")
        #��������� ����:
        #������ ������������ ����� �� ����������� � ����������
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

    #������� ����� ������ ����������	
    try { 
        $sap = ((invoke-WebRequest $webReq -ContentType "text/plain; charset=utf-8" -UseBasicParsing).content | convertFrom-Json)
    } catch {
        #�������!
        $err=$_.Exception.Response.StatusCode.Value__
        Log("WARNING: user ["+$user.sAMAccountname+"] with Name ["+$user.displayName+"] not found in SAP by $($webReq)")
        #��������� �� ����� �:
        #� ����� ���������� ����� ���� ���� ���������. ���� ��� �� ���������� �� �����:
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

#��������� ������������
function ParseUser() {
	param
	(
		[object]$user
	)
	#���������� ������
	#��������� ����������� � �� �� ����
	$needUpdate = $false
	#��������������� ������������ � �� �� ����
	$needRename = $false
	#��������� ������������ � �� �� ����
	$needDismiss = $false

	
	$sap = FindUser($user)
	#���� ������������ �������
	if ( -not ($sap -eq "error")) {
		
		if ($sap.Uvolen -eq "1") {
			#������� ����� ������
			if ($sap.resign_date.Length -gt 0) {
				$resign_date=[datetime]::parseexact($sap.resign_date, $dateformat_SAP, $null)
				#��� ������?
				if ((Get-Date) -gt $resign_date) {
					$needDismiss = $true
				}
			}
		}
		if ($needDismiss) {
			#��������� ���������
			#��������� ���������� ���������
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
			$title=$sap.Doljnost
			if ($title.Length -gt 128) {
				#����������� ����� ����
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

            #��������� �����
			if ($user.EmployeeNumber -ne $sap.employee_id ){
				Log($user.sAMAccountname+": got Numbr ["+$user.EmployeeNumber+"] instead of ["+$sap.employee_id+"]")
				$user.EmployeeNumber=$sap.employee_id
				$needUpdate = $true
			}
			
            #ID �����������
            if ($mutiorg_support) {
			    if ($user.EmployeeID -ne $sap.org_id ){
				    Log($user.sAMAccountname+": got orgID ["+$user.EmployeeID+"] instead of ["+$sap.org_id+"]")
				    $user.EmployeeID=$sap.org_id
				    $needUpdate = $true
			    }
            }


            if ($mobile_from_SAP) {
			    #$correctedMobile=correctMobile($sap.Mobile)
			    #��������� ������������� ������, �.�. �� ������ � ����� ��������� ������� �����
			    #�� ������ ��� �������� ��������� � ������ ���� ��� � ��
			    #� ������� ��� ����������� �� ������
			    $correctedMobile= correctPhonesList($sap.Mobile)
			    if ([string]$user.mobile -ne [string]$correctedMobile) {
				    #��� ���� ���������� ������ ��������� �� ������ ���� ��� ����� ������, �.�. ��� �������� ��������
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
