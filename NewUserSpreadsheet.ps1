########Notes##########
<#
This is a mostly function script to pull users from a .CSV file
and add them to Active Directory and Exchange.

Basic user/mailbox creation works.

With little work you can have a working script.

Feel free to modify to quite your environment.

#>
###########Variables############
#Modify according to your environment.
$mailBoxDB = 'MailboxDB'
$exchange = 'ExchangeServer1'
$Company = "Aceme Dildo Company"

#Semi random password for setup.
$number = Get-Random -Maximum 9999 -Minimum 1000
$password = "Welcome" + $number
$password += (33..47)| Get-Random -Count 2 | % {[char]$_}
$password = $password.replace(' ','')
#Make password string usable.
$SecPaswd= ConvertTo-SecureString -String $password –AsPlainText –Force 

$domain = "@corp.adc.com"
$Global:ou = 'ADC Users'#default users OU
#Add your standard groups to a list
$group1= "Allusers"
#$group2 = "Group 2"
#$group3 = "Group 3"
#Make it a collection for easy use
#$groups = $group1,$group2,$group3
$groups = $group1

$date = [datetime]::Today.ToString('MM-dd-yyyy')

$Global:fName = $null
$Global:lName = $null
$Global:email = $null
$global:uname = $null
$Global:title= $null
$Global:department = $null
######## End Variables ############

#ToDo Figure out file path to be same dir as script
#This needs to be modified, reads all users but only performs action on the last user read in.
Import-Csv -path C:\temp\NewEmployee.csv | foreach {
    $fName = $_.FirstName
    $lName = $_.LastName
    $email = $_.email
    $title = $_.title
    $department = $_.department
    $ou = $_.office

    #create username
    $uname = $lName + $fName.substring(0,1)}
    #Simple cleanup
    $uname= $uname.replace("'",'')

function Show-User{
    #Write-Host $uname
    #Check for blanks - if x is blank prompt to fill in
    #create username
    Write-Host "****Summary****"
    Write-Host "Employee name: $fName $lName"
    Write-Host "Email Address: $email"
    Write-Host "Username: $uname"
    Write-Host "Title: $title"
    Write-Host "Department: $department"
    Write-Host "Office: $ou"

    
}
function YouGood{
     $uGood = Read-Host -Prompt "Does this look good? Y/N"
    if ($uGood -eq 'Y'){
        #Proceed with user creation
        Write-Host 'Proceeding with User Creation.'
        CreateUser
    }
    else {
    #if you need correction call CorrectUser function
        CorrectUser
    }   

}

function CorrectUser
{
    cls
    do{
        #anything else is a fail which means correction needed
       
        #Test for blank input
        #function to test for null?
        Write-Host "1.First name: $fName"
        Write-Host "2.Last name: $lName"
        Write-Host "3.Username: $uname"
        Write-Host "4.Email Address: $email"
        Write-Host "5.Title: $title"
        Write-Host "6.Department: $department"
        Write-Host "7.Office: $ou"
        
        Write-Host "Enter 0 when you're done making corrections."
        $pCorrect = Read-Host -Prompt "What do you need to correct?" 
        
        switch ($pCorrect)
        {
             '1' {
                 $up = Read-Host -Prompt "Enter new First Name."
                 $Global:fName = $up
                 $Global:uname = $lName + $fName.substring(0,1)
                 $Global:email = $fname + '.' + $lName + "@adc.com"
                cls
             } '2' {
                $up = Read-Host -Prompt "Enter new Last Name."
                 $Global:lName = $up
                 $Global:uname = $lName + $fName.substring(0,1)
                 $Global:email = $fname + '.' + $lName + "@adc.com"
                cls
             } '3' {
                 $up = Read-Host -Prompt "Enter new username."
                 $Global:uname = $up
         cls
             }'4'{
                $up = Read-Host "Enter new email."
                $Global:email = $up
            cls   
             } '5' {
                    $up = Read-Host -Prompt "Enter new title."
                $Global:title = $up
               cls
                 
             } '6' {
                $up = Read-Host -Prompt "Enter new department."
                $Global:department = $up
               cls
             } '7'{
                $up= Read-host -Prompt "Enter new office."
                $Global:ou = $up
               cls
             }
         
        }

        }
        until ($pCorrect -eq '0')
        cls
        #Show-User
}



function CreateUser{
    Write-host "We're going to create a new user now."
    $Name = $Global:fName + " " + $Global:lName

    #Office Determines OU
    switch ($Global:OU)
    {
         'Field' {
             $Global:OU = 'OU=Field,DC=corp,DC=adc,DC=com'
         } 'Office' {
             $Global:OU = 'OU=Office,DC=corp,DC=adc,DC=com'
         } 'Remote' {
             $Global:OU = 'OU=Remote,DC=corp,DC=adc,DC=com'
         }'' {
             #if no option is entered leave OU as is.
         }
    }

    Write-Host $Global:OU

    #Establish Connections to Exchange and AD - Will prompt for admin credentials
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchange/PowerShell/ -Authentication Kerberos -Credential $UserCredential
    Import-PSSession $Session
    Import-Module ActiveDirectory

    #Use New-Mailbox so a mailbox is created along with user.
    #Remove -Whatif to actually create user
    New-Mailbox –Name $Name –Alias $global:uname –OrganizationalUnit $Global:ou -Database $mailBoxDB –UserPrincipalName ($global:uname +$domain) -FirstName $global:fName –LastName $global:lName –ResetPasswordOnNextLogon $true -Password $SecPaswd -whatif

    ##Pause to allow user completeion to complete.
    ##If you try to modify a new user too quickly, sometiems fails.
    Write-Host "Pausing to allow user creation to complete."
    Start-Sleep -s 5
    
    #user existing user for testing
    #again, remote -Whatif to complete command
    Set-ADUser $uname -Title $Global:title -Department $Global:department -WhatIf

    Write-Host "User $uname was created with a password of $password."
}

function addGroupMembership {
    #this function is to update the user's group membership based on role and office. or Maybe just office and the other stuff can be done later.
    
    #user existing user for testing
    ##UPDATE THIS VARIABLE
    $user = 'usert'
    Set-ADUser $user -Title $Global:title -Department $Global:department -WhatIf
   

}

#This do statement really ties the script together.
#Needs additional work since it will only take the last user in the .CSV file
do
{
    Show-User
     $uGood = Read-Host -Prompt "Does this look good? Y/N"
     switch ($uGood)
     {
           'N' {
                CorrectUser   
                
           } 'Y' {
                cls
                #Write-Host "$fName $lName"
                CreateUser
                #clear variables
                #quit
                #Remove-PSSession $Session
                Clear-Variable my* -Scope Global
                return
           } '3' {
                cls
               passwordReset
           } '4'{
                cls
                defaultPassword
           } 'q' {
                #Do some cleanup
                Remove-PSSession $Session
                Clear-Variable my* -Scope Global
                return
           }
     }
}
until ($input -eq 'q')
