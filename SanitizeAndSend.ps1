<#
This script will sanitize, ie replace strings, so you can upload
a clean version to github or wherever.

Second half can zip it and all relevant files and email to you
using a local exchange server.
#>

#Path to script you want to cleanup
$path = "$PSScriptRoot\DirtyScript.ps1"
 
#Creates a new file
((Get-Content -path $path -Raw) -replace 'yourdomain.com','adc.com') | Set-Content -Path "$PSScriptRoot\CleanScript.ps1"

#If you have additional strings to replace, no need to create new file 
((Get-Content -path "$PSScriptRoot\CleanScript.ps1" -Raw) -replace 'Company Name','Acme Dildo Company') | Set-Content -Path "$PSScriptRoot\CleanScript.ps1"

#Specify script file and all other files in Path
$compress = @{
	Path = "$PSScriptRoot\CleanScript.ps1", "C:\temp\input.txt"
	CompressionLevel = "Fastest"
	DestinationPath = "$PSScriptRoot\CleanScript.zip"
}

#Compress into a .zip archive for emailing
Compress-Archive @compress

#Email settings
$smtpServer="mail.company.com"
$from = "Test User test.user@adc.com"
$subject="Sanitized Powershell Script"
$emailaddress= your.email@adc.com
$body = "Powershell script with company references removed."

#Send it
Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML  -Encoding $textEncoding -Attachments "$PSScriptRoot\CleanScript.zip" 
