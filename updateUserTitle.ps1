#import-module activedirectory

$userlist = import-csv C:\users\jackson.li\Desktop\ToDo\UpdateTitle.csv

foreach ($user in $userlist) {

$newUser = $user.name.replace(" ",".")

Set-ADUser $newUser -Replace @{title=$user.Title} -server cnszpdc004.premiumsoundsolutions.com

}