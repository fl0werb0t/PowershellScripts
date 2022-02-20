#Open Exel doc
#get user details
#format
#verify before creating
#option to modify something?
#confirm you want to create
#RandomPassword generate


###########Variables############
#Modify according to your environment.
$mailBoxDB = 'Database'
$exchange = 'ExchangeServer'
$Company = "Acme Dildo Company"
#$terminatedOU = 'OU=Terminated Employees,DC=corp,DC=adc,DC=com' #Terminated Users OU
#'Semi random password for setup.
$password = 'Welcome' #Generic password for setup.
$num = Get-Random -Minimum 1000 -Maximum 9999
#Make password string usable.
$SecPaswd= ConvertTo-SecureString -String $password –AsPlainText –Force 

$domain = "@corp.adc.com"
$OU = 'Forerunner users'#default users OU
#Add your standard groups to a list
$group1= "Allusers"
#$group2 = "Group 2"
#$group3 = "Group 3"
#Make it a collection for easy use
#$groups = $group1,$group2,$group3
$groups = $group1

$date = [datetime]::Today.ToString('MM-dd-yyyy')

$Global:fName
$Global:lName
$Global:email
$Global:title
$Global:department
$Global:ou


######## End Variables ############

#Figure out file path to be same dir as script
Import-Csv -path C:\drivers\NewEmployee.csv | foreach {
    $fName = $_.FirstName
    $lName = $_.LastName
    $email = $_.email
    $title = $_.title
    $department = $_.department
    $ou = $_.office

    #create username
    $uname = $lName + $fName.substring(0,1)}

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

    do{
        #anything else is a fail which means correction needed
        Write-host "Something needs correction. Pick your option."
        Write-Host "Enter 0 when you're done making corrections."
        
        #Test for blank input
        #function to test for null?

        $pCorrect = Read-Host -Prompt "What do you need to correct?"
        switch ($pCorrect)
        {
             '1' {
                 $up = Read-Host -Prompt "Enter new First Name."
                 $fName = $up
                 $uname = $lName + $fName.substring(0,1)
                 $email = $fname + '.' + $lName + "@frtinc.com"
                 Show-User
             } '2' {
                $up = Read-Host -Prompt "Enter new Last Name."
                 $lName = $up
                 $uname = $lName + $fName.substring(0,1)
                 $email = $fname + '.' + $lName + "@frtinc.com"
                 Show-User
             } '3' {
                 $up = Read-Host -Prompt "Enter new username."
                 $uname = $up
                 Show-User
             }'4'{
               $up = Read-Host -Prompt "Enter new email."
                 $email = $up
                 Show-User
             } '5' {
                 $up = Read-Host -Prompt "Enter new department."
                 $department = $up
                 Show-User
                 $pCorrect = '0'
             } '6' {
               #other option
             }
         
        }

        }
        until ($pCorrect -eq '0')
        cls
        Show-User
}

function CreateUser{
    Write-host "We're going to create a new user now."
    # New-Mailbox –Name $Name –Alias $Username –OrganizationalUnit $OU -Database $mailBoxDB –UserPrincipalName $UPN -FirstName $fName –LastName $lName –ResetPasswordOnNextLogon $true -Password $SecPaswd

    #if yes than proceed with user creation
    #if it's not right pick what you want
    #this could probably be a switch
    #this might need to be in a loop
    Write-Host "User $uname was created."
   
}

#This do statement really ties the script together.
do
{
    #user is already pulled in at initial run
    Show-User
     $uGood = Read-Host -Prompt "Does this look good? Y/N"
     switch ($uGood)
     {
           'N' {
                CorrectUser   
                
           } 'Y' {
                cls
                Write-Host "$fName $lName"
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


# match column name ie $_.department
#$givenName = $_.name.split()[0] 
#$surname = $_.name.split()[1]
#new-aduser -name $_.name -enabled $true –givenName $givenName –surname $surname -accountpassword (convertto-securestring $_.password -asplaintext -force) -changepasswordatlogon $true -samaccountname $_.samaccountname –userprincipalname ($_.samaccountname+”@ad.contoso.com”) -city $_.city -department $_.department

Show-User