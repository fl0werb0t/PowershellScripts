##Automate a few steps when setting up a new Windows machine.
##Modify to fit your needs

#ToDo Start Menu setup
#ToDo Sign in to O365?

#Create a local admin user
$password = 'Password123' #Generic password for setup.
#Make password string usable.
$SecPaswd= ConvertTo-SecureString -String $password –AsPlainText –Force 
New-LocalUser "newUser" -Password $SecPaswd -FullName "New User"
#Add them to local admins 
Add-LocalGroupMember -Group Administrators -Member newUser

#Set password as expired, user prompted to change next login
$euser = [ADSI]"WinNT://localhost/newUser,user"
$euser.PasswordExpired = 1
$euser.setInfo()

Write-Host "Enter computer name."
$Nombre = Read-Host
Rename-Computer $Nombre
Write-Host "Restart to apply."

#Disabled Dell services
$Services = get-service "*Dell*"
foreach ($s in $Services)
{
    Stop-Service $s.ServiceName
    Set-Service $s.ServiceName -StartupType  Disabled
}

#Disable Microsoft services
Set-Service "dcpm-notify" -StartupType Disabled
Stop-Service "dcpm-notify" -Force

Set-Service "DDVDataCollector" -StartupType Disabled
Stop-Service "DDVDataCollector" -Force

Set-Service "DDVRulesProcessor" -StartupType Disabled
Stop-Service "DDVRulesProcessor" -Force

Set-Service "DDVCollectorSvcApi" -StartupType Disabled
Stop-Service "DDVCollectorSvcApi"-Force

Set-Service "SupportAssistAgent" -StartupType Disabled
Stop-Service "SupportAssistAgent" -Force

#Download and install programs silently
#Install file silently as Current User
#Uses Chrome as an example
$Installer = "$env:temp\chrome_installer.exe"
$url = 'http://dl.google.com/chrome/install/375.126/chrome_installer.exe'
Invoke-WebRequest -Uri $url -OutFile $Installer -UseBasicParsing
try {
    Start-Process -FilePath $Installer -Args '/silent /install' -Wait
    Remove-Item -Path $Installer
} catch {
    # $_ returns the error details
    Write-Host "Installer returned the following error $_"
}


#Use this one if you need to run as a differnt user
<#

$c = Get-Credential

$Installer = "$env:temp\chrome_installer.exe"
$url = 'http://dl.google.com/chrome/install/375.126/chrome_installer.exe'
Invoke-WebRequest -Uri $url -OutFile $Installer -UseBasicParsing
try {
    Start-Process -FilePath $Installer -Args '/silent /install' -Wait -Credential $c
    Remove-Item -Path $Installer
} catch {
    # $_ returns the error details
    Write-Host "Installer returned the following error $_"
}
#>
