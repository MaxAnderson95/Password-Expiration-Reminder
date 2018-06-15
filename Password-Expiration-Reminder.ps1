#### User Defined Variables ####

$SearchBaseOU = "" #OU in DN format
$DaysToRemind = 1,3,5 #How many days before expiration to remind. This accepts multiple days in an array.
$LogFileName = "" #Path to log file for output

$EmailFrom = "Network Manager <NetworkManager@contoso.com>" #Email address the emails will come from
$EmailBCC = "" #Email address of an administrator to monitor if the emails were sent successfully
$EmailSubject = "Reminder: Your Network/Email Password is Expiring Soon"
$EmailSMTPServer = "" #SMTP Server to send email via

$ChangePasswordHelpURL = "" #Link to article with instructions on chaning password

#### User Defined Variables ####

#### Helper Functions ####

Function DaysFromNow ($Days) { 
    
    (Get-Date).AddDays($Days).Date

}

Function Write-Log {

    Param (
      
      [String]$Message,
      
      [String]$LogFile
      
    )
    
    $Output = "$(Get-Date -Format G): $Message"
    
    Add-Content -Value $Output -Path $LogFile
  
}

#### Helper Functions ####

#### Core Script ####

<#
Get a list of users, who are enabled, whose passwords are not set to never expire, whose passwords are not set to change on next logon, and who have an email address
from the Users OU. Then take the UserPasswordExpiryTime and convert it to a datetime object
#>
$Users = Get-ADUser -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and PasswordExpired -eq $False -and PasswordLastSet -ne 0 -and EmailAddress -ne 0} `
 -Properties "Name", "EmailAddress", "msDS-UserPasswordExpiryTimeComputed", "PasswordExpired" `
 -SearchBase $SearchBaseOU `
 | Select-Object -Property "Name","GivenName","EmailAddress", `
    @{Name = "PasswordExpirationDay"; Expression = {([datetime]::FromFileTime($_.'msDS-UserPasswordExpiryTimeComputed')).Date}} `
 | Where-Object {$_.PasswordExpirationDay -gt (Get-Date)}


Write-Log -Message "Starting Script" -LogFile $LogFileName

ForEach ($User in $Users) {

    Write-Log -Message "Processing $($User.Name)" -LogFile $LogFileName

    ForEach ($Day in $DaysToRemind) {

        If (($User.PasswordExpirationDay) -eq (DaysFromNow($Day))) {
            
            If ($Day -gt 1) {
                
                $DayPluralOrSingular = "days"
            
            } Else {
                
                $DayPluralOrSingular = "day"
            
            }

            $EmailParamerters = @{

                To = $User.EmailAddress
                From = $EmailFrom
                BCC = $EmailBCC
                Subject = $EmailBCC
                SMTPServer = $EmailSMTPServer
                Body = @"
Hello $($User.givenname),<br>
<br>
Your network/email password is set to expire in <b>$Day $DayPluralOrSingular</b>. It is <b>HIGHLY</b> recommended that you change your password <b>BEFORE</b> it expires to prevent any interuption to work.
<br>
<br>
Go to this page to learn how to change your password: <a href="$ChangePasswordHelpURL">$ChangePasswordHelpURL</a>
<br>
<br>
If you have any trouble changing your password, open a ticket by emailing <a href="mailto:help@contoso.com">help@contoso.com</a>
<br>
<br>
<i>Note this email was sent automatically. Please do not reply to this email, you will not recieve a response.</i>
"@
                BodyAsHTML = $True

            }

            Try {
                
                Write-Log -Message "Sending Email to $($User.Name)" -LogFile $LogFileName
                Send-MailMessage @EmailParamerters

            }

            Catch {

                Write-Log -Message "There was an error sending an email to $($User.Name)" -LogFile $LogFileName
                Write-Log -Message $_ -LogFile $LogFileName

            }

        }

    }

}

Write-Log -Message "Script Complete" -LogFile $LogFileName

#### Core Script ####
