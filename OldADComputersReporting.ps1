import-module activedirectory 
$logdate = Get-Date -format yyyyMMdd
$subjectdate = Get-Date -format "MM/dd/yyyy"
$logfile = "F:\Scripts\Reporting\logs\ExpiredComputers\ExpiredComputers - "+$logdate+".csv"
$mail = "reporting-techops@corp.example.com"
$smtpserver = "smtp-relay.gmail.com"
$emailFrom = "reporting-techops@corp.example.com"
$domain = "10.239.0.146"
$emailTo = "$mail"
$subject = "Old Computers in Active Directory" + " " + $subjectdate
$DaysInactive = 30
$time = (Get-Date).Adddays(-($DaysInactive))
$body = 
    "Attached is a list containing Computers inactive 30 days or more.

    Tech Ops Reporting"
 
# Change this line to the specific OU that you want to search
$searchOU = "DC=AD,DC=corp.example,DC=COM"

# Get all AD computers with LastLogon less than our time
Get-ADComputer -SearchBase $searchOU -Filter {LastLogon -lt $time -and enabled -eq $true} -Properties LastLogon, description|
 
# Output hostname and LastLogon into CSV
select-object Name,DistinguishedName, description, enabled,@{Name="Stamp"; Expression={[DateTime]::FromFileTime($_.LastLogon)}} | export-csv $logfile -notypeinformation

Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -Attachments $logfile -SmtpServer $smtpserver