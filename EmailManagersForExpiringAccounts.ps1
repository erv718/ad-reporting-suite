# Get enabled user accounts that are expiring within the next 33 days
$aboutToExpireUsers = Search-ADAccount -AccountExpiring `
    -DateTime (Get-Date).AddDays(33).Date `
    -UsersOnly |
    Where-Object { $_.Enabled }

# Retrieve manager and expiration date for each enabled, expiring user
$UserInfo = $aboutToExpireUsers |
    ForEach-Object {
        Get-ADUser -Identity $_.SamAccountName `
            -Properties manager, AccountExpirationDate, Enabled
    } |
    Where-Object { $_.Enabled }

# Group by manager
$managers = $UserInfo | Group-Object -Property manager

foreach ($manager in $managers.Name | Select-Object -Unique) {
    # Get manager’s email
    $managerMail = (Get-ADUser -Identity $manager -Properties mail).mail
    # All direct reports under this manager
    $reports = $UserInfo | Where-Object { $_.manager -eq $manager }

    $count = $reports.Count
    $accountVerbiage = if ($count -eq 1) { "account" } else { "accounts" }
    $isAreVerbiage   = if ($count -eq 1) { "is" } else { "are" }

    # HTML table styling
    $Header = @"
<style>
  TABLE {border: 1px solid black; border-collapse: collapse;}
  TH {text-align: center; border: 1px solid black; padding: 3px; background-color: #F0F8FF;}
  TD {text-align: center; border: 1px solid black; padding: 3px;}
</style>
"@

    # Build email body
    $body = @"
Hello,<br><br>
The following network $accountVerbiage $isAreVerbiage near expiration:<br><br>
$(
    $reports |
    Select-Object `
      @{Name="Username";Expression={$_.SamAccountName}}, `
      @{Name="Name";Expression={$_.Name}}, `
      @{Name="Expiration Date";Expression={($_.AccountExpirationDate).ToShortDateString()}} |
    ConvertTo-HTML -Head $Header
)<br>
As their manager, please confirm if they will continue with their role.<br><br>
<b>If they need an extension, please reply to this email with your approval and we can extend up to 1 year.</b><br><br>
Thank You,<br><br>
[Company] Information Technology
"@

    # Send notification
    Send-MailMessage `
      -From       "techcommunications@corp.example.com" `
      -To         "helpdesk@corp.example.com" `
      -Cc         $managerMail `
      -Subject    "[ACTION REQUIRED] - Accounts expiring in the next 30 days!" `
      -Body       $body `
      -BodyAsHtml $true `
      -SmtpServer "smtp-relay.gmail.com"
}
