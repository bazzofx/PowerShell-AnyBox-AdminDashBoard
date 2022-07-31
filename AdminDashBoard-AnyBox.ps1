$userProfile = $env:USERPROFILE
$dataPath ="$userProfile\Downloads\UsersToDisable.csv"
$path ="$userProfile\Downloads\report.csv"
$count=1
$users = ""
$credential = "your-email-here"
$license = "x_NAME_OF_YOUR_LICENSE_HERE_x"
$ErrorActionPreference ="Stop"


#########################################
#########################################
#############VARIABLES ABOVE ############



#------MINOR Functions
function loadCSV{
    Try{
    $users = Import-CSV $box.search -Header "UserPrincipalName", "Login", "Password"
    $global:users = $users
    #$users = Get-Content $box.search |Select -skip 1 | ConvertFrom-CSV -Header "UserPrincipalName", "Login", "Password"
    Write-Host "New Data loaded.." -ForegroundColor White -BackgroundColor Green}
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
installDependencies
#-------------------------------------

#-----MAIN Functions
function login-AzureAD{
Show-AnyBox -Message "You will be asked to login into MicrosfotAzure AD" -Buttons "OK"

    Try{
    Connect-AzureAD
    Write-Host "You are now connected to Microsoft Azure AD" -ForegroundColor Green}
    Catch [System.Exception] {Write-Warning $Error[0]}
} #-- close function
function login-365{
Show-AnyBox -Message "You will be asked to login into Microsfot Online 365" -Buttons "OK"

    Try{
    Connect-MsolService
    Write-Host "You are now connected to Microsoft Online 365" -ForegroundColor Green}
    Catch [System.Exception] {Write-Warning $Error[0]}
    Catch {"Failed to connecte to server, please check if you are connected on the VPN"}
} #-- close function

function resetPassAzure{
login-AzureAD  
loadCSV


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
     }
     Catch [System.Exception] {Write-Warning $Error[0]}
     Catch {   Write-Host "[ERROR] : $currentRowEmail PASSWORD FAIL TO UPDATED." -ForegroundColor Red     }
     
     $count +=1   
        } # --close ForLoop
        

} #-- close function 
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
        

} #-- close function
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
        

} #-- close function

function lockAccount {
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
        }
        Catch [System.Exception] {Write-Warning $Error[0]
                                  Write-Host $currentRowEmail -ForegroundColor Red}
        Catch {Write-Host "[ERROR] - Something went while trying to BLOCK $currentRowEmail" -ForegroundColor Red}
        $count += 1
        }
Show-AnyBox -Message "Accounts have been blocked!" -Buttons "OK"
} #-- close function
function unlockAccount {
loadCSV
Show-AnyBox -Message "Accounts are now unlocked" -Buttons "OK"
    
    
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
        }
        Catch [System.Exception] {Write-Warning $Error[0]
        Write-Host $currentRowEmail -ForegroundColor Red}
        Catch {Write-Host "[ERROR] - Something went wrong while UNBLOCKING $currentRowEmail" -ForegroundColor Red}
        $count += 1
        }

} #-- close function

function addLicense {
loadCSV

    
    
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
   }
Catch [System.Exception] {Write-Warning $Error[0]}
Catch {Write-Host "[ERROR] - Something went when attemption to add '$license' license to $currentRowEmail" -ForegroundColor Red}
    $count+=1
    } #-close forLoop

Show-AnyBox -Message "License '$license' added to accounts." -Buttons "OK"
        } #-- close function
function removeAllLicenses {
loadCSV
    
    
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
        Write-Host "[SUCCESS] - The Licenses from $email has been removed." -ForegroundColor Green
        }
   }
Catch [System.Exception] {Write-Warning $Error[0]}
Catch {Write-Host "[ERROR] - Something went wrong while removing licenses from $email" -ForegroundColor Red}
       $count +=1
} # -- close ForEach
   
        } #--close function

         
# --                                   POP UP SETTINGS
#-------------------------VARIABLES
$p = @(New-AnyBoxPrompt -InputType 'FileOpen' -Name "search" -Message 'Open File:' -ReadOnly)
$p += @(New-AnyBoxPrompt -InputType Text -Name "check" -Message "Onboarding" -ValidateSet 'Reset Password','Unlock Account','Add License' -ShowSetAs Radio)
$p += @(New-AnyBoxPrompt -InputType Text -name "check2" -Message "Offboarding" -ValidateSet 'Lock Account','Remove All Licenses' -ShowSetAs Radio)


$b = @(New-Button -Text "Login" -Onclick {login-365})
$b += @(New-Button -Text "Execute")
$b += @(New-Button -Text "Cancel")
$m = "Users Admin Dashboard v1.0"
# --                                   POP UP SETTINGS


#--------->>>>> OPEN DIALOG BOX  <<<<<<<<<<<--------
$box = Show-AnyBox -Prompt $p -Message $m -Buttons $b

#-----------------------------------------------


#----- Logic POP UP BOX

if($box.check -eq "Reset Password") {resetPassAzure}
elseif($box.check -eq "Unlock Account") {unlockAccount}
elseif($box.check -eq "Add License") {addLicense}
elseif($box.check2 -eq "Lock Account"){lockAccount}
elseif($box.check2 -eq "Remove All Licenses"){removeAllLicenses}
else {Write-Warning "Please make a selection !"}
