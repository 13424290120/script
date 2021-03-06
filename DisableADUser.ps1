##########################################################################################
#
# Script Name: DisableADUser.ps1
# Description: Disable the AD user and Office365 user account
# Author: Jackson Li
# Date: 20th July 2017
# Command Syntax: C:\Users\jackson.li\Desktop\ToDo\disableADuser.ps1 -ADuser brolin.luo -shareTo none
#
##########################################################################################
param(
[string]$ADUser=$(throw "Parameter missing: -ADUser UserName; Syntax: disableADuser.ps1 -ADuser firstName.lastName -shareTo user/none") ,
[string]$shareTo=$(throw "Parameter missing: -shareTo UserName; Syntax: disableADuser.ps1 -ADuser username -shareTo user/none")
)



Function msg($message)
{
    Write-Host $message
}

Function sendMail($recipient, $user)
{
    [string]$mailHost = "mailhost.premiumsoundsolutions.com"
    [string]$currentDate = Get-Date
    Send-MailMessage -From noreply@premiumsoundsolutions.com -To $recipient `
     -Subject "$user account has been disabled" -SmtpServer $mailHost `
     -Body "The user $user has been disabled from Active Directory and Office365 at $currentDate "
}

# Import PowerShell Module
Function importModule($module)
{
    Import-Module $module    
    if($?){
        msg("$module  Module imported!!! Please wait a moment")
    }else{
        msg("$module import failed, please install RSAT from https://www.microsoft.com/en-us/download/details.aspx?id=7887")
        break
    }
    
}

# Disable Active Directory User Account
Function disableADAccount($username)
{
    set-ADuser $username -enable 0 -clear Department,Manager -replace @{msExchHideFromAddressLists=$true} -Server "cnszpdc004.premiumsoundsolutions.com"
    if($?){
        msg("$username has been disabled in the Active Directory!!!")
    }
}

# Connect to Office365 Powershell Management Platform
Function connectO365
{
    $UserCredential = Get-Credential
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
    if(!$?){
        msg("Sorry, connect to Office365 failed!!!")
        break
    }else{
        Import-PSSession $Session
    }
}

# Disable Office365 User and convert to shared mailbox, then remove the license
Function convertToShareMailBox($O365User)
{
    Set-Mailbox $O365User -Type Shared
    if($?){
        msg("$O365User mailbox has been convert to Shared mailbox")
    }
#    AzureAD module needed/PowerShell 3.0/4.0    
#    Set-MsolUserLicense -UserPrincipalName -join($O365User,"premiumsoundsolutions.com") -RemoveLicenses "premiumsoundsolutions:ENTERPRISEPACK"
#    Set-MsolUser -UserPrincipalName fabricec@litwareinc.com -BlockCredential $true
}

# Delegate the mailbox access to other user
Function delegateMailbox([string]$initiator, [string]$delegateTo)
{

    Add-MailboxPermission -Identity $initiator -User $delegateTo -AccessRights FullAccess -InheritanceType All
    if($?){
        msg("$initiator mailbox has been shared to $delegateTo")
    }
}

# Restore the deleted user mailbox
Function restoreMailbox($deletedUser)
{
    $fullMailbox = -join($deletedUser,"@premiumsoundsolutions.com")
    Restore-MsolUser -UserPrincipalName $fullMailbox
    return $?
}


# Main Program

$executionPolicy = Get-ExecutionPolicy
$smICT = "jackson.li@premiumsoundsolutions.com"


if($executionPolicy -eq "RemoteSigned")
{    


    try{importModule("ActiveDirectory")}
    catch{Write-Warning $_}

    try{disableADAccount($ADUser)}
    catch{Write-Warning $_}

    try{connectO365}
    catch{Write-Warning $_}

    try{convertToShareMailBox($ADUser)}
    catch{Write-Warning $_}
    
    if( $shareto -ne "none" )
    {
        try{delegateMailbox -initiator $ADUser -delegateTo $shareTo}
        catch{Write-Warning $_}        
    }
    
    try{sendMail -recipient $smICT -user $ADUser}
    catch{Write-Warning $_}
    
}else{
$warningMessage=@"    
    Sorry, please change your powershell execution policy as RemoteSigned first.
    the command is "Set-ExecuitionPolicy RemoteSigned".
    You have to run this command as administrator.
"@
    Write-Host $warningMessage
            
}

