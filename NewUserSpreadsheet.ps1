########Notes##########
<#
This is a mostly function script to pull users from an Excel file
and add them to Active Directory and Exchange.

Basic user/mailbox creation works.

With little work you can have a working script.

Feel free to modify to suite your environment.

#>

####ToDo Copy User access from existing User#####

###########User Variables############
#Modify according to your environment.
$mailBoxDB = 'Database'
$exchange = 'ExchangeServer'
$Company = "Acme Dildo Company."
 
#Specify domain and default user OU
$domain = "@corp.adc.com"
$Global:ou = 'Acme Dildo Company Users'#default users OU

#Add your standard groups to a list
$group1= "Allusers"
#Make it a collection for easy use
#$groups = $group1,$group2,$group3
$groups = $group1
########End User Variables ###########

##### Global Variables #####
#Semi random password for setup.
$number = Get-Random -Maximum 9999 -Minimum 1000
$password = "Welcome" + $number
#Make password string usable.
$SecPaswd= ConvertTo-SecureString -String $password –AsPlainText –Force

#Declare global variables
$Global:fName = $null
$Global:lName = $null
$Global:email = $null
$global:uname = $null
$Global:title= $null
$Global:department = $null
$Global:Manager = $null


######## Create Connections #######
 #Establish Connections to Exchange and AD - Will prompt for admin credentials
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$exchange/PowerShell/ -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session
Import-Module ActiveDirectory

##Read in all users into an array
#Add a warning that file is open, optino to close?

$path = "c:\drivers\newuser.xlsx"
#Cannot read file if it is open.
#Check if file is locked, give user a chance to close, otherwise exit.
try { 
    [IO.File]::OpenWrite($path).close();
    $true 
    #Write-Host "File is closed and able to be read."
}
catch {
        #false if file is locked
    $false
    Write-Host "File is open. Please close to continue. We'll wait."
    sleep -Seconds 15
    try { 
    [IO.File]::OpenWrite($path).close();
    $true 
    #Write-Host "File is closed and able to be read."
    }
    catch {
        #false if file is locked
        #script will exit if file is still locked
    $false
    Write-Host "File is still open. Please close and run script again."
    }
}

import-module psexcel #it wasn't auto loading on my machine

$people = new-object System.Collections.ArrayList

foreach ($person in (Import-XLSX -Path $path -RowStart 1))
{
$people.add($person) | out-null #I don't want to see the output
}


function Show-User{
    #Write-Host $uname
    #Check for blanks - if x is blank prompt to fill in
    #create username
    Write-Host "****Summary****"
    Write-Host "Employee name: $Global:fName $Global:lName"
    Write-Host "Email Address: $email"
    Write-Host "Username: $Global:uname"
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

function generatePassword{
    $number = Get-Random -Maximum 9999 -Minimum 1000
    $password = "Welcome" + $number
    #Make password string usable.
    $SecPaswd= ConvertTo-SecureString -String $password –AsPlainText –Force
}

function CreateUser{
    Write-host "We're going to create a new user now."
    $Name = $Global:fName + " " + $Global:lName

  #Office Determines OU
  switch ($Global:OU)
    {
         '1' {
             $Global:OU = 'OU=Field,DC=corp,DC=adc,DC=com'
         } '2' {
             $Global:OU = 'OU=Users,OU=Office,DC=corp,DC=adc,DC=com'
         }  '' {
             #if no option is entered leave OU as is.
         }
    }

    Write-Host $Global:OU

   

    #Use New-Mailbox so a mailbox is created along with user.
    #Remove -Whatif to actually create user
    New-Mailbox –Name $Name –Alias $global:uname –OrganizationalUnit $Global:ou -Database $mailBoxDB –UserPrincipalName ($global:uname +$domain) -FirstName $global:fName –LastName $global:lName –ResetPasswordOnNextLogon $true -Password $SecPaswd -Whatif

    ##Pause to allow user completeion to complete.
    ##If you try to modify a new user too quickly, sometiems fails.
    Write-Host "Pausing to allow user creation to complete."
    Start-Sleep -s 3
    
    #user existing user for testing
    #again, remote -Whatif to complete command
    #If you get an Insufficient access rights error here, make sure you run this script from an Admin powershell.

    #Set-ADUser $Global:uname -Title $Global:title -Department $Global:department -Manager $Global:Manager -WhatIf
    #Add-ADGroupMember -Identity 'Allusers' -Members $Global:uname -whatif


    Write-Host "User $uname was created with a password of $password."
}

function addGroupMembership {
    #this function is to update the user's group membership based Department? or Maybe just office and the other stuff can be done later.
    
    #user existing user for testing
    ##UPDATE THIS VARIABLE
    $user = 'usert'
  
   

}

#This makes ties it all together.
Foreach($person in $people){
    #first populate variables for current user
    $Global:fName = $person.firstname
    $Global:lName = $person.lastname
    $Global:email = $person.email
    $global:uname = $person.username
    $Global:title = $person.title
    $Global:department = $person.department
    $Global:OU = $person.office

    
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
                    break
               } '3' {
                    cls
                   passwordReset
               } '4'{
                    cls
                    defaultPassword
               } 'q' {
                    #Do some cleanup
                    Remove-PSSession $Session
                    Remove-Variable -Name * -ErrorAction SilentlyContinue
                    return
               }
         }
 }

#Do some cleanup
Remove-PSSession $Session
Remove-Variable -Name * -ErrorAction SilentlyContinue





