import-module ActiveDirectory
set-ADuser "amy.zhang" -enable 0
set-ADuser "amy.zhang" -replace @{msExchHideFromAddressLists=$true}
set-ADuser "amy.zhang" -clear Department
set-ADuser "amy.zhang" -clear Manager

Set-ExecutionPolicy RemoteSigned
Get-Module -ListAvailable |foreach {"`r`nmodule name: $_"; "`r`n";gcm -Module $_.name -CommandType cmdlet, function | select name}
Install-Module -Name AzureAD 

$UserCredential = Get-Credential

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

set-mailbox amy.zhang -Type Shared
Add-MailboxPermission -Identity "Amy Zhang" -User henry.zhang -AccessRights FullAccess -InheritanceType All
Set-MsolUserLicense -UserPrincipalName belindan@litwareinc.com -RemoveLicenses "litwareinc:ENTERPRISEPACK"
Restore-MsolUser -UserPrincipalName BelindaN@litwareinc.com



$UserCredential = Get-Credential
Connect-MsolService -Credential $UserCredential
