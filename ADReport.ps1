.\PSHTML-AD.ps1

$logdate = Get-Date -format yyyyMMdd
$subjectdate = Get-Date -format "MM/dd/yyyy"
$logfile = "F:\Scripts\Reporting\logs\WeeklyADReports\WeeklyADReports - "+$logdate+".csv"
$mail = "reporting-techops@corp.example.com"
$smtpserver = "smtp-relay.gmail.com"
$emailFrom = "reporting-techops@corp.example.com"
$emailTo = "$mail"
$subject = "Weekly AD Report" + " " + $subjectdate
$time = (Get-Date).Adddays(-($DaysInactive))
$body = $htmlbody + $htmlbody2 + $htmlbody3




Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Attachments $logfile