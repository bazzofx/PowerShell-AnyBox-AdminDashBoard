$userProfile = $env:USERPROFILE
$dataPath ="$userProfile\Downloads\UsersToDisable.csv"
$path ="$userProfile\Downloads\report.csv"
$count=1
$users = ""
$credential = "your-email-here"
$license = "YOUR_LICENSE_HERE"
$ErrorActionPreference ="Stop"
##################################################################
#--------------LOG VARIABLES START -------------------------------
$logPath = "$userProfile\Documents\Reports\AdminDashBoard\"
$logGoodFile = "SUCCESS $(get-date -f dd-MM-yyy).log"
$logBadFile = "EXCEPTIONS $(get-date -f dd-MM-yyy).log"
$logGood = $logPath + $logGoodFile
$logBad = $logPath + $logBadFile
$checkPathlogGood = Test-Path $logGood
$checkPathlogBad = Test-Path $logBad
$testFolderPath = Test-Path $logPath

Try{
New-Item -Path $logPath -ItemType Directory
Write-Host "All log files will be saved $logPath" -foreground Yellow}
Catch{Write-Host "All log files will be saved $logPath" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "" -ForegroundColor Yellow
Write-Host $logPath -ForegroundColor Yellow
}

if ($checkPathlogGood -eq $false -and $checkPathlogBad -eq $false){
$newFileGood = New-Item $logGood
$newFileBad = New-Item $logBad
}
Else{
<#rm $logGood
rm $logBad
$newFileGood = New-Item $logGood
$newFileBad = New-Item $logBad #>
} 
#---------------------LOG VARIABLES END-------------------------------
######################################################################
#


#-----------NOT BEING USED-------------
function resetPassAD{
loadCSV
    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        $securePassword = ConvertTo-SecureString $currentRowPassword -AsPlainText -Force



     Try{
     Set-ADAccountPassword -Identity $currentRowEmail -NewPassword $securePassword -Reset
     Write-Host " The password for $currentRowEmail has been reset to ---> |  $currentRowPassword" -ForegroundColor Green
     Write-Host ""
     Write-Host "---------------------------"


     }
     Catch [System.Exception] {Write-Warning $Error[0]}
     Catch {   Write-Host "[ERROR] : $currentRowEmail PASSWORD FAIL TO UPDATED." -ForegroundColor Red     }
     
     $count +=1   
        } # --close ForLoop
        

} #-- close function - legacy not being used
function resetPass365{
loadCSV
    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        $securePassword = ConvertTo-SecureString $currentRowPassword -AsPlainText -Force



     Try{
     Set-MsOlUserPassword -UserPrincipalName $currentRowEmail -NewPassword $securePassword -ForceChangePassword $False
     Write-Host " The password for $currentRowEmail has been reset to ---> |  $currentRowPassword" -ForegroundColor Green
     Write-Host ""
     Write-Host "---------------------------"
     }
     Catch [System.Exception] {Write-Warning $Error[0]}
     Catch {   Write-Host "[ERROR] : $currentRowEmail PASSWORD FAIL TO UPDATED." -ForegroundColor Red     }
     
     $count +=1   
        } # --close ForLoop
        

} #-- close function - legacy not being used
#-----------NOT BEING USED-------------



#------------MINOR FUNCTIONS-------------
function loadCSV{
    Try{
    $users = Import-CSV $box.search -Header "UserPrincipalName", "Login", "Password"
    $global:users = $users
    #$users = Get-Content $box.search |Select -skip 1 | ConvertFrom-CSV -Header "UserPrincipalName", "Login", "Password"
    Write-Host "             CSV Data Loaded correctly.           " -ForegroundColor White -BackgroundColor Green}
    Catch{Write-Host "Failed to import data."}
                  }
function installDependencies {
if (Get-Module -ListAvailable -Name AnyBox) {
Import-Module AnyBox
} 
else {
Install-Module AnyBox
Import-Module Anybox
Write-Host "Installing Dependencies" -ForegroundColor Yellow
} 
}
function Start-log{
Add-Content $logGood "------------>                Starting Log File on               $(Get-Date)"
Add-Content $logBad "------------>                Starting Log File on               $(Get-Date)"
}
function checkConnection{
    Get-MsolDomain -ErrorAction SilentlyContinue
    if($?){ 
    Write-Host "You are already logged in to MsOline Service 365" -ForegroundColor Green}
    else{ Write-Host "You will need to login before you use this application" -ForegroundColor Yellow}
                        }
installDependencies
checkConnection
Start-Log
#----------------------------------------
#--------------MAIN FUNCTIONS------------

function login-AzureAD{
Show-AnyBox -Message "You will be asked to login into MicrosfotAzure AD" -Buttons "OK"

    Try{
    Connect-AzureAD
    Write-Host "You are now connected to Microsoft Azure AD" -ForegroundColor Green
                Add-Content $logGood "User Login to Azure AD successfully"}
    Catch [System.Exception] {Write-Warning $Error[0]
                Add-Content $logBad $Error[0]
                Add-Content $logBad "Failed to login into Azure"
                Add-Content $logBad "-------------------------------"
    }
} #-- close function
function login-365{
Show-AnyBox -Message "You will be asked to login into Microsfot Online 365" -Buttons "OK"

    Try{
    Connect-MsolService
    Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green
                Add-Content $logGood "Login to Microsoft Online 365 successfull"}
    Catch [System.Exception] {Write-Warning $Error[0]
                Add-Content $logBad $Error[0]
                Add-Content $logBad "Microsfot Online 365 was not able to connect"
                Add-Content $logBad "-------------------------------"}
    Catch {"Failed to connecte to server, please check if you are connected on the VPN"
                Add-Content $logBad "Failed to connecte to server, please check if you are connected on the VPN"
                Add-Content $logBad $Error[0]
                Add-Content $logBad "-------------------------------"
    }
} #-- close function

function resetPassAzure{
login-AzureAD  
loadCSV

Add-Content $logGood "----------- PASSWORD RESET SUCCESS LOG START --------------"
Add-Content $logBad "------------ PASSWORD RESET EXCEPTION LOG START --------------" 

    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        $securePassword = ConvertTo-SecureString $currentRowPassword -AsPlainText -Force



     Try{
    $user = Get-AzureADUser -Filter "UserPrincipalName eq 'susan.tiley@fitzroy.org'"
    Set-AzureADUserPassword -ObjectId $currentRowEmail -Password $securepassword -ForceChangePasswordNextLogin $true  
     
     Write-Host " The password for $currentRowEmail has been reset to ---> |  $currentRowPassword" -ForegroundColor Green
     Write-Host ""
     Write-Host "---------------------------"

                 Add-Content $logGood " The password for $currentRowEmail has been reset to ---> |  $currentRowPassword"
     }#-close Try
     Catch [System.Exception] {Write-Warning $Error[0]
                Add-Content $logBad $Error[0]
                Add-Content $logBad "[ERROR] $currentRowEmail PASSWORD FAIL TO UPDATED."}
     Catch {   Write-Host "[ERROR] : $currentRowEmail PASSWORD FAIL TO UPDATED." -ForegroundColor Red     
                Add-Content $logBad $Error[0]
                Add-Content $logBad "[ERROR] $currentRowEmail PASSWORD FAIL TO UPDATED."}
     
     $count +=1 
  
        } # --close ForLoop
                Show-AnyBox -Message "Accounts on your list had PASSWORD UPDATED, please check exceptions. " -Buttons "OK"
                Add-Content $logGood "----------- PASSWORD RESET SUCCESS LOG END --------------"
                Add-Content $logGood ""
                Add-Content $logBad "------------ PASSWORD RESET EXCEPTION LOG END --------------"
                Add-Content $logBad ""         

} #-- close function 
function lockAccount {
Add-Content $logGood "----------- BLOCKING ACCOUNT SUCCESS LOG START --------------"
Add-Content $logGood ""
Add-Content $logBad "----------- BLOCKING ACCOUNT EXCEPTION LOG START --------------"
Add-Content $logBad ""
loadCSV
    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        
        Try{
        Set-Msoluser -UserPrincipalName $currentRowEmail -BlockCredential $true
        Write-Host "[SUCCESS] - $currentRowEmail has been BLOCKED" -ForegroundColor yellow
        Write-Host ""
        Write-Host "---------------------------"

        Add-Content $logGood "[SUCCESS] - $currentRowEmail has been BLOCKED"
        }
        Catch [System.Exception] {Write-Warning $Error[0]
                                  Write-Host $currentRowEmail -ForegroundColor Red
                                  Add-Content $logBad $Error[0]
                                  Add-Content $logBad "[ERROR] - Something went while trying to BLOCK $currentRowEmail"
                                  }
        Catch {Write-Host "[ERROR] - Something went while trying to BLOCK $currentRowEmail" -ForegroundColor Red
                                  Add-Content $logBad $Error[0]
                                  Add-Content $logBad "[ERROR] - Something went while trying to BLOCK $currentRowEmail"  }
        $count += 1
        }#- Close ForLoop
Show-AnyBox -Message "Accounts on your list are now BLOCKED, please check exceptions. " -Buttons "OK"
        Add-Content $logGood "----------- BLOCKING ACCOUNT SUCCESS LOG ENDS --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- BLOCKING ACCOUNT EXCEPTION LOG ENDS --------------"
        Add-Content $logBad ""
 
} #-- close function
function unlockAccount {
loadCSV

Add-Content $logGood "----------- UNLOCKING ACCOUNT SUCCESS LOG START --------------"
Add-Content $logBad "----------- UNLOCKING ACCOUNT EXCEPTION LOG START --------------"    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        
        Try{
        Set-Msoluser -UserPrincipalName $currentRowEmail -BlockCredential $false
        Write-Host "[SUCCESS] - $currentRowEmail has been UNLOCKED" -ForegroundColor Green
        Write-Host ""
        Write-Host "---------------------------"

        Add-Content $logGood "[SUCCESS] - $currentRowEmail has been UNLOCKED"
        }
        Catch [System.Exception] {Write-Warning $Error[0]
        Write-Host $currentRowEmail -ForegroundColor Red
        Add-Content $logBad $Error[0]
        Add-Content $logBad "[ERROR] - Something went wrong while UNLOCKING $currentRowEmail"
        }
        Catch {Write-Host "[ERROR] - Something went wrong while UNLOCKING $currentRowEmail" -ForegroundColor Red
        Add-Content $logBad "[ERROR] - Something went wrong while UNLOCKING $currentRowEmail"        
        }
        $count += 1

        } # -Close ForLoop
Show-AnyBox -Message "Accounts on your list have been UNLOCKED, please check exceptions. " -Buttons "OK"
        Add-Content $logGood "----------- UNLOCKING ACCOUNT SUCCESS LOG ENDS --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- UNLOCKING ACCOUNT EXCEPTION LOG ENDS --------------"
        Add-Content $logBad ""
 
} #-- close function
function addLicense {
loadCSV
        Add-Content $logGood "----------- ADDING LICENSES SUCCESS LOG START --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- ADDING LICENSES EXCEPTION LOG START --------------"
        Add-Content $logBad ""
    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        $license = $global:license

        Try{
     Set-MsolUserLicense -UserPrincipalName $currentRowEmail -AddLicenses $license 
     Write-Host "[SUCCESS] - License $license added to $currentRowEmail" -ForegroundColor Green
     Write-Host ""
     Write-Host "---------------------------"

     Add-Content $logGood "[SUCCESS] - License $license added to $currentRowEmail"
   }
Catch [System.Exception] {Write-Warning $Error[0]

    Add-Content $logBad $Error[0]
    Add-Content $logBad "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail"}
Catch {Write-Host "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail" -ForegroundColor Red

    Add-Content $logBad "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail"
}
    $count+=1

    } #-close forLoop

Show-AnyBox -Message "Accounts on your list were given licensen $license, please check exceptions. " -Buttons "OK"
        Add-Content $logGood "----------- ADDING LICENSES SUCCESS LOG END --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- ADDING LICENSES EXCEPTION LOG END --------------"
        Add-Content $logBad ""

        } #-- close function
function removeAllLicenses {
loadCSV
        Add-Content $logGood "----------- REMOVAL ALL LICENSES SUCCESS LOG START --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- REMOVAL ALL LICENSES EXCEPTION LOG START --------------"
        Add-Content $logBad ""    
    
    ForEach ($row in $global:users)
        {
        $currentRowLogin = $global:users.Login[$count]
        $currentRowPassword = $global:users.Password[$count]
        $currentRowEmail = $global:users.UserPrincipalName[$count]
        $license = "FitzRoy:STANDARDWOFFPACK"
Try{        
    (get-MsolUser -UserPrincipalName $currentRowEmail).licenses.AccountSkuId |
    foreach{
        Set-MsolUserLicense -UserPrincipalName $currentRowEmail -RemoveLicenses $_
        Write-Host "[SUCCESS] - The Licenses from $currentRowEmail has been removed." -ForegroundColor Green

        Add-Content $logGood "[SUCCESS] - The Licenses from $currentRowEmail has been removed."

        }
   }
Catch [System.Exception] {Write-Warning $Error[0]
        Add-Content $logBad $Error[0]
        Add-Content $logBad "[ERROR] - Something went wrong while removing licenses from $currentRowEmail"
}
Catch {Write-Host "[ERROR] - Something went wrong while removing licenses from $email" -ForegroundColor Red
        Add-Content $logBad $Error[0]
        Add-Content $logBad "[ERROR] - Something went wrong while removing licenses from $currentRowEmail"
}
       $count +=1
} # -- close ForEach

Show-AnyBox -Message "Accounts on your list had ALL licensed REMOVED, please check exceptions. " -Buttons "OK"
        Add-Content $logGood "----------- REMOVAL ALL LICENSES SUCCESS LOG END --------------"
        Add-Content $logGood ""
        Add-Content $logBad "----------- REMOVAL ALL LICENSES EXCEPTION LOG END --------------"
        Add-Content $logBad "" 
 
        } #--close function

         
# --                                   MAIN WINDOW POP UP SETTINGS
#--------------------VARIABLES-------------------
$p = @(New-AnyBoxPrompt -InputType 'FileOpen' -Name "search" -Message 'Open File:' -ReadOnly)
$p += @(New-AnyBoxPrompt -InputType Text -Name "check" -Message "Onboarding" -ValidateSet 'Reset Password','Unlock Account','Add License' -ShowSetAs Radio)
$p += @(New-AnyBoxPrompt -InputType Text -name "check2" -Message "Offboarding" -ValidateSet 'Lock Account','Remove All Licenses' -ShowSetAs Radio)


$b = @(New-Button -Text "Login" -Onclick {login-365})
$b += @(New-Button -Text "Execute")
$b += @(New-Button -Text "Cancel")
$m = "Users Admin Dashboard v1.0"

#--------->>>>> OPEN DIALOG BOX  <<<<<<<<<<<--------
$Window = Show-AnyBox -Prompt $p -Message $m -Buttons $b 
#-----------------------------------------------


#------------- Logic MAIN POP UP BOX ----------------

if($Window.check -eq "Reset Password") {resetPassAzure}
elseif($Window.check -eq "Unlock Account") {unlockAccount}
elseif($Window.check -eq "Add License") {addLicense}
elseif($Window.check2 -eq "Lock Account"){lockAccount}
elseif($Window.check2 -eq "Remove All Licenses"){removeAllLicenses}
else {Write-Warning "Please make a selection !"}
