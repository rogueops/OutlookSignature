#import the active directory module which is needed for Get-ADUser
import-module activedirectory

$path = "C:\sigs\"
If(!(test-path -PathType container $path))
{
      New-Item -ItemType Directory -Path $path
}

#set folder location for files, the folder must already exist
$save_location = 'c:\sigs\'

<#EXAMPLE
$users = Get-ADUser -filter * -searchbase "OU=Employees,OU=Users,DC=MyDomain,DC=com" -Properties * -Credential domain\DAdmin -Server domain.com
#>
$users = Get-ADUser -filter {(EmailAddress -like "*") -and (Enabled -eq "True")} -searchbase "OU=SomeOU,DC=YourDomain,DC=com" -Properties *

#Fields to pull info from
foreach ($user in $users) {
$account_name = "$($User.sAMAccountName)"
$full_name = "$($user.GivenName) $($User.sn)$($User.extensionAttribute1)"
$job_title = "$($User.title)"
$comp = "$($User.company)"
$phone = "$($User.telephoneNumber)"
$mobile = "$($User.mobile)"
$fax = "$($User.facsimileTelephoneNumber)"
$email = "$($User.emailaddress)"
$street = "$($User.streetAddress)"
$city = "$($User.l)"
$state = "$($User.st)"
$zipcode = "$($User.postalCode)"
$mobileyes = "$($Null)"
$faxyes = "$($Null)"
$mobileAndFax = "$($Null)"
$noMobileOrFax = "$($Null)"

#Webpage field for logo from web page
$logo = "$($User.wWWHomePage)"

if ($mobile)
{$mobileyes = "1"}
else {}
if ($fax)
{$faxyes = "2"}
else {}
if ($mobileyes+$faxyes -eq "12")
{$mobileAndFax = "3"}
else {}
if ($mobileyes+$faxyes -eq "")
{$noMobileOrFax = "4"}
else {}

if ($MobileAndFax -eq "3")
{
#We need to construct and write the html signature file
$output_file = $save_location + $account_name + ".htm"
Write-Host "Now attempting to create signature html file for " $full_name
"<span style=`"font-family: calibri,sans-serif;`"><strong>" + $full_name + "</strong><br />", $job_title + "<br />", "<a href=`"http://www.website.com`" target=`"_blank`">$comp</a>" + "<br />", $street + "<br />", $city + ", " + $state + ", " + $zipcode + "<br />", "<span style=`"font-weight: bold;`">" + "Ph:" + "</span>" + " " + $phone + "<br />", "<span style=`"font-weight: bold;`">" + "Cell:" + "</span>" + " " + $mobile + "<br />", "<span style=`"font-weight: bold;`">" + "Fax:" + "</span>" + " " + $fax + "<br />", "<a href=`"mailto: $Email`">$Email</a>" + "<br />", "<img alt=`"Corporate Logo`" border=`"0`" height=`"108`" src=`"" + $logo + "`" width=`"173`" />", "</span><br />" | Out-File $output_file
}
elseif ($mobileYes -eq "1")
{
#We need to construct and write the html signature file
$output_file = $save_location + $account_name + ".htm"
Write-Host "Now attempting to create signature html file for " $full_name
"<span style=`"font-family: calibri,sans-serif;`"><strong>" + $full_name + "</strong><br />", $job_title + "<br />", "<a href=`"http://www.website.com`" target=`"_blank`">$comp</a>" + "<br />", $street + "<br />", $city + ", " + $state + ", " + $zipcode + "<br />", "<span style=`"font-weight: bold;`">" + "Ph:" + "</span>" + " " + $phone + "<br />", "<span style=`"font-weight: bold;`">" + "Cell:" + "</span>" + " " + $mobile + "<br />", "<a href=`"mailto: $Email`">$Email</a>" + "<br />", "<img alt=`"Corporate Logo`" border=`"0`" height=`"108`" src=`"" + $logo + "`" width=`"173`" />", "</span><br />" | Out-File $output_file
}
elseif ($faxyes -eq "2")
{
#We need to construct and write the html signature file
$output_file = $save_location + $account_name + ".htm"
Write-Host "Now attempting to create signature html file for " $full_name
"<span style=`"font-family: calibri,sans-serif;`"><strong>" + $full_name + "</strong><br />", $job_title + "<br />", "<a href=`"http://www.website.com`" target=`"_blank`">$comp</a>" + "<br />", $street + "<br />", $city + ", " + $state + ", " + $zipcode + "<br />", "<span style=`"font-weight: bold;`">" + "Ph:" + "</span>" + " " + $phone + "<br />", "<span style=`"font-weight: bold;`">" + "Fax:" + "</span>" + " " + $fax + "<br />", "<a href=`"mailto: $Email`">$Email</a>" + "<br />", "<img alt=`"Corporate Logo`" border=`"0`" height=`"108`" src=`"" + $logo + "`" width=`"173`" />", "</span><br />" | Out-File $output_file
}
elseif ($NoMobileOrFax -eq "4")
{
#We need to construct and write the html signature file
$output_file = $save_location + $account_name + ".htm"
Write-Host "Now attempting to create signature html file for " $full_name
"<span style=`"font-family: calibri,sans-serif;`"><strong>" + $full_name + "</strong><br />", $job_title + "<br />", "<a href=`"http://www.website.com`" target=`"_blank`">$comp</a>" + "<br />", $street + "<br />", $city + ", " + $state + ", " + $zipcode + "<br />", "<span style=`"font-weight: bold;`">" + "Ph:" + "</span>" + " " + $phone + "<br />", "<a href=`"mailto: $Email`">$Email</a>" + "</span><br />" | Out-File $output_file
}
else {}
}
