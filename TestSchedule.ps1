<#

### To schedule a powershell script ###

program/script

C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

add arguments

-NonInteractive -WindowStyle Hidden -command D:\move\movemailbox.ps1

#>

Send-MailMessage -From noreply@premiumsoundsolutions.com -To jackson.li@premiumsoundsolutions.com `
-Subject "Test Schedule Powershell Script" -SmtpServer mailhost.premiumsoundsolutions.com `
-Body "Test message only" 