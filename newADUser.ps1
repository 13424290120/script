$xl=New-Object -ComObject excel.application
$xl.visible=$true
$wb=$xl.workbooks.open("C:\Users\jackson.li\Desktop\ToDo\NewUserRequestList.xlsx")
$sheet=$wb.sheets.item(1)
$object=
$addressList="\\cnszpst002\general\it\addadaccount\address.csv"

#for($i=1;$i -lt $sheet.range("a65536").end(-4162).row;$i++) #xlDown -4121 Down. xlToLeft -4159 To left. xlToRight -4161 To right. xlUp -4162 Up.
#{

Function removePoundSign($argText){
    $newText = $argText.split('#')
    return $newText[-1]
}

Function combineName($argName){
    $newName = $argName.replace(" ",".")
    return $newName
}

$i = 2

$Site = removePoundSign($sheet.range("A$i").text) #A列对应每行的值 A column correspond to per row value
$onBoardDate = $sheet.range("B$i").text
$firstName = $sheet.range("C$i").text
$lastName = $sheet.range("D$i").text
$department = removePoundSign($sheet.range("E$i").text)
$businessCategory = removePoundSign($sheet.range("F$i").text)
$jobTitle =  $sheet.range("G$i").text
$telephoneNumber =  $sheet.range("H$i").text
$mobileNumber = $sheet.range("I$i").text
$manager = combineName($sheet.range("J$i").text)
$PSSDomain = "@premiumsoundsolutions.com"
$accountName = -join($firstName,".",$lastName)
$principalName = -join($accountName,$PSSDomain)
$displayName = -join($firstName," ",$lastName)
#}

$useraddr=(import-csv -path $addressList | where-object {$_.location -eq $Site})

#Import-Module ActiveDirectory

Write-host new-ADuser $accountName -givenName $firstName -surName $lastName `
-office $Site -userPrincipalName $principalName -samAccountName $accountName -displayname $displayName `
-description "PSS - AP" -officephone $telephoneNumber -emailaddress $principalName -homepage "www.PremiumSoundSolutions.com" `
-city $Site -country "CN"  -scriptpath "logon" -homedrive "H:" `
-homedirectory -join("\\cnszpst002\users$\",$accountName) -mobilephone $MobileNumber `
-company "Premium Sound Solutions" -title $userobject.title -department $department -manager $manager `
-streetaddress $useraddr.officeaddress -pobox $useraddr.officepobox -postalcode $useraddr.officezippostalcode -state $useraddr.state `
-otherattributes @{'proxyaddresses'=("SMTP:" + $accountName + $PSSDomain),("smtp:" + $accountName + "@dmpss.com");'businesscategory'=$businessCategory;'co'="China";'countrycode'="156";'employeeType'="Employee"} `
-accountpassword (convertto-securestring "Welcome@pss" -asplaintext -force) -Enable $True -path $useraddr.ou

 
$xl.displayAlerts=$False
$wb.Close()
$xl.Application.Quit()

