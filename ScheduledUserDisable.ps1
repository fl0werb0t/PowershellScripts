##################################
<#
Use this simple script to disable an AD user.
It will forward their emails to a specified user
and send an email upon completion.

Use with Task Scheduler. 

https://github.com/fl0werb0t/

#>


##################################

#Variables
$Username = 'johnUser' #User to disable
$ForwardtoEmail = 'person@company.com' #email to forward to.
$terminatedOU = 'OU=Terminated Employees,DC=corp,DC=company,DC=com' #Terminated Users OU
$emailaddress= @("manager@company.com") #Who you want notified, can be multiple emails
$from = "Help Desk <helpdesk@company.com>" 
$subject="User Disabled"

######### Email settings ######
$textEncoding = [System.Text.Encoding]::UTF8
$exchange = 'ExchangeServer' #Exchange server
$smtpServer="mail.company.com" #SMTP address

try{
#Import AD Module so we can do AD stuff.
Import-Module ActiveDirectory

#Disable the account, get display name and move to terminated users OU
Disable-ADAccount -Identity $Username
$name = get-aduser $Username -Properties * |Select-Object -ExpandProperty Name
Get-ADUser $Username | Move-ADObject -TargetPath $terminatedOU

#Update the body.
$body = "This is an automated message. </br>"
$body +="User $name  has been disabled."

#Connect to Exchange to handle email forwarding.
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchange/PowerShell/
Import-PSSession $Session

Set-Mailbox -Identity $Username -ForwardingAddress $ForwardtoEmail -HiddenFromAddressListsEnabled $true -DeliverToMailboxAndForward $true

#Update the body again
$body += "</br> Their email will be forwarded to $ForwardtoEmail."

#Send email.
Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML  -Encoding $textEncoding   

#Be nice and clean up.
Clear-Variable my* -Scope Global
}
Catch
{
    #Send failure message.
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    $emailaddress = "helpdesk@company.com"
    Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $ErrorMessage -bodyasHTML  -Encoding $textEncoding  
    Break
}