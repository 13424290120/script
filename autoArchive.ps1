<#   
This script compress all klippel log files in the current folder and make new .7zip file. 
  
Author: Jackson Li

Date: 2017-07-25

Any problem please contact jackson.li@premiumsoundsolutions.com
  
#> 
 
 
#### 7 zip variable I got it from the below link  
 
# Alias for 7-zip 
if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) {throw "$env:ProgramFiles\7-Zip\7z.exe needed"} 
set-alias sz "$env:ProgramFiles\7-Zip\7z.exe" 
 
############################################ 
#### Variables  

$productionLine = "BW1"
$destArchiveFolder = "C:\Test\Dest"
$KlippelFilePath = "C:\Test\Source"
$localSQLBackup = "$KlippelFilePath\SQLBackup"
$databaseName = "EOL_Generic"
$DBBackupFile = "$localSQLBackup\$databaseName.bak"
$sysAdmin = "jackson.li@premiumsoundsolutions.com"
$productionEngineer = "jackson.li@premiumsoundsolutions.com"
$mailHost = "mailhost.premiumsoundsolutions.com"
$currentDate = get-Date -Format yyyyMMdd
$currentYear = get-Date -Format yyyy
$storageLocation = "$destArchiveFolder\$productionLine\$currentYear\$currentDate"

########### END of VARABLES ################## 

################ FUNCTIONS ################## 

function getCurrentFolder($UNC)
{
    $arrFolder = $UNC.split("\")
    return $arrFolder[-1]
}


function createFolder($path)
{

    if (-not (Test-Path "$path")){
        try
        {
            New-Item -ItemType directory -Path "$path" -ErrorAction stop
        }
        catch
        {
            Write-Verbose -Verbose $Error[0]
        }
        
    }

}

############ END OF FUNCTIONS ############## 

# To create the local SQL database backup folder.
createFolder($localSQLBackup)

# To create the backup folder on remote storage system ex:EOL1\2017\20171011\xxxxxxx.
createFolder($storageLocation)

# To import sqlps powershell module

try
{
    Import-Module sqlps -ErrorAction stop
    Backup-SqlDatabase -ServerInstance '.\SQLExpress' -Database $databaseName -BackupFile $DBBackupFile -Verbose
    
}
catch
{
$errorMessage =
@"
Import Powershell Module sqlps failed!

Please install the following packages in this order and then restart computer:

Microsoft® System CLR Types for Microsoft® SQL Server® 2012 (SQLSysClrTypes.msi)
Microsoft® SQL Server® 2012 Shared Management Objects (SharedManagementObjects.msi)
Microsoft® Windows PowerShell Extensions for Microsoft® SQL Server® 2012 (PowerShellTools.msi)
Be sure to select the appropriate package platform for each, either x86 or x64.
"@

Write-Host $errorMessage
}

# Get all klippel test data folders

$klippelFolders = Get-ChildItem -Path $KlippelFilePath | Where-Object { $_.Mode -match "^d" } 

foreach ($folder in $klippelFolders){
    $folderName = $folder.name
    #$archiveLog = "ArchiveLog_"+$folderName+"_"+$currentDate+".txt"
    $archiveFile = $folderName+".7z" 
    
    # Compress the kdbx file into 7z file
    
    sz a -t7z "$KlippelFilePath\$archiveFile" "$KlippelFilePath\$folderName\" -sdel
    
    # Create the folder again to avoid impact to klippel   
    New-Item -ItemType directory -Path $KlippelFilePath\$folderName
}

### Move the 7z and log file to the Archive location ###

$archiveFiles = Get-ChildItem -Path $KlippelFilePath | Where-Object { $_.Extension -eq ".7z" }

foreach ($zfile in $archiveFiles) {
    $zFileName = $zFile.name 
    $zFileDirectory = $zFile.DirectoryName
    
    Copy-Item "$zFileDirectory\$zFileName" "$storageLocation"
    if ($?){
        Remove-Item "$zFileDirectory\$zFileName"
        Send-MailMessage -From noreply@premiumsoundsolutions.com -To $productionEngineer `
         -Subject "$productionLine archive klippel test data succeed!!!" -SmtpServer $mailHost `
         -Body "The klippel test data $zFileDirectory\$zFileName has been archive into $destArchiveFolder\$productionLine"
    }else{
        Send-MailMessage -From noreply@premiumsoundsolutions.com -To $productionEngineer `
         -Subject "$productionLine archive klippel test data failed!!!" -SmtpServer $mailHost `
         -Body "If you don't know why it's failed, please contact $sysAdmin"    
    }
}
               
########### END OF SCRIPT ########## 
