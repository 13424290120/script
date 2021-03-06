<#
# Script Name : sendScoreCard.ps1
# Description : Send the score cards from sharepoint to respective suppliers
# Author : Jackson Li
# Date : 14th Aug 2017
# Support : jackson.li@premiumsoundsolutions.com
#>

### Variables ###

$supplierListFile = "C:\Users\jackson.li\Desktop\ToDo\supplierListTest.csv"
$scoreCardFolder =  "C:\Users\jackson.li\Desktop\ToDo\scoreCards"
$mailHost = "mailhost.premiumsoundsolutions.com"
$ccList = @(
"ellen.zhang@premiumsoundsolutions.com",
"patrick.han@premiumsoundsolutions.com",
"robbie.zhang@premiumsoundsolutions.com",
"roger.xie@premiumsoundsolutions.com",
"sy.xie@premiumsoundsolutions.com",
"terry.xiao@premiumsoundsolutions.com",
"tony.zhang@premiumsoundsolutions.com",
"wilson.luo@premiumsoundsolutions.com",
"yifan.wang@premiumsoundsolutions.com",
"owen.liu@premiumsoundsolutions.com")
$message = @"
<p>Dear Supplier:</p>
<p>Please kingly find attached latest GSRS.</p>
<p>If you have any query, feel free to contact  me.</p>
Thanks!
"@


### Function ###

Function getSupplierName($supplierName){
$originalSupplierName = $supplierName

$originalSupplierName -replace "[]"

}

### Loading Score Cards file ###
if (-not (test-path $scoreCardFolder)){

    throw("ERROR: The scoreCard files not found, please check the scoreCard file path.")

}else{

    ### Loading Supplier List ###
    
    if (-not (test-path $supplierListFile)){

        throw("The supplier file is not found, please check the supplier list file location!")

    }else{

        $supplierList = Import-Csv $supplierListFile

    }

           
    foreach ($supplier in $supplierList){
            
        $supplierName = $supplier.SupplierName
        $supplierEmail = $supplier.SupplierEmail
        $PssContact = $supplier.PssContact
        
        #Write-Host $supplierName
        
        $scoreCards = @(Get-ChildItem $scoreCardFolder | Where-object { $_.name -Match $supplierName -and $_.name -Match "APAC.pdf$"} | sort-object LastWriteTime -desc)
        
        if ($scoreCards.count -ge 1){
        
            [string]$mailAttachment = $scoreCardFolder+"\"+$scoreCards[0].Name
        
            #Write-Host "Debug:" $scoreCards[0].Name + "|" + $supplierEmail + "|" + $PssManager + "|" + $PssContact + "|" + $mailAttachment
            
            Send-MailMessage -To $supplierEmail -Subject "PSS Supplier ScoreCard" -Body $message -Attachments $mailAttachment -SmtpServer $mailHost -BodyAsHtml -From $PssContact -Cc $ccList
            
            if ($?){
                Write-Host $scoreCards[0] "has been sent to" $supplierEmail
            }else{
                Write-Host "Failed to Send" $scoreCards[0] to $supplierEmail
            }
            
        }
            
    }
    
}




