# Define HTML Styling
$head = @"
<style>
    @charset "UTF-8";
    table { font-family: "Trebuchet MS", Arial, Helvetica, sans-serif; border-collapse: collapse; }
    td { font-size: 1em; border: 1px solid #33adff; padding: 5px; }
    th { font-size: 1.1em; text-align: center; padding: 5px; background-color: #4db8ff; color: #ffffff; }
    tr { background-color: #b3e0ff; } /* Removed incorrect selector */
</style>
"@

# Define Timeframe for Report (Last 7 Days)
$When = ((Get-Date).AddDays(-7)).Date

# Define Excluded Groups
$excludedGroups = @(
    "All Club Computers", "FD Club Computers", "MF Club Computers", "BO Club Computers",
    "Kiosk Club Computers", "sec00us-contractors-sec", "sec00us-googleuser-sec",
    "Club FD&MF Web Filter"
)

# Function to Get Object Creator from Active Directory
Function Get-Creator($distinguishedName) {
    try {
        $object = Get-ADObject -Identity $distinguishedName -Properties msDS-CreatorSID -ErrorAction Stop
        if ($object -and $object.'msDS-CreatorSID') {
            return (New-Object System.Security.Principal.SecurityIdentifier($object.'msDS-CreatorSID')).Translate([System.Security.Principal.NTAccount]).Value
        }
    }
    catch {
        return "Unknown"
    }
}

# Fetch New Users Created in Last 7 Days
$users = Get-ADUser -Filter {whenCreated -ge $When} -Properties Created, Description, PasswordExpired, DistinguishedName |
    Select-Object Name, Created, SamAccountName, Description, PasswordExpired, DistinguishedName, @{Name="CreatedBy";Expression={Get-Creator $_.DistinguishedName}} |
    Sort-Object Created -Descending

# Fetch New Groups Created in Last 7 Days
$groups = Get-ADGroup -Filter {whenCreated -ge $When} -Properties Created, DistinguishedName |
    Select-Object Name, Created, DistinguishedName, @{Name="CreatedBy";Expression={Get-Creator $_.DistinguishedName}}

# Fetch New Computers Created in Last 7 Days
$comps = Get-ADComputer -Filter {Created -ge $When} -Properties Created, OperatingSystem, MemberOf, DistinguishedName |
    Select-Object Name, Created, DistinguishedName, OperatingSystem,
        @{Name="SecurityGroups";Expression={($_.MemberOf | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "}},
        @{Name="CreatedBy";Expression={Get-Creator $_.DistinguishedName}} |
    Sort-Object Created -Descending

# Fetch Disabled Users (Including OU Information)
$disabledUsers = Get-ADUser -Filter {Enabled -eq $false} -Properties mail, Modified, DistinguishedName |
    Where-Object { $_.Modified -ge $When } |
    Select-Object Name, SamAccountName, mail, Modified,
        @{Name="OU";Expression={($_.DistinguishedName -split ",")[1] -replace "OU=",""}},
        @{Name="DisabledBy";Expression={Get-Creator $_.DistinguishedName}}

# Track Group Membership Changes
$groupChanges = @()
$allGroups = Get-ADGroup -Filter * -Properties Members, whenChanged

foreach ($group in $allGroups) {
    if (($group.Name -notin $excludedGroups) -and ($group.whenChanged -ge $When)) {
        $members = $group.Members | ForEach-Object { Get-ADObject -Identity $_ -Properties Name, DistinguishedName }
        foreach ($member in $members) {
            $groupChanges += [PSCustomObject]@{
                GroupName = $group.Name
                MemberName = $member.Name
                MemberDistinguishedName = $member.DistinguishedName
                ChangeDate = $group.whenChanged
                ModifiedBy = Get-Creator $group.DistinguishedName
            }
        }
    }
}

# Convert Data to HTML
$htmlSections = @(
    $users  | ConvertTo-HTML -Head $head -PreContent "<H3>New Users - Last 7 Days</H3>" | Out-String
    $groups | ConvertTo-HTML -Head $head -PreContent "<H3>New Groups - Last 7 Days</H3>" | Out-String
    $comps  | ConvertTo-HTML -Head $head -PreContent "<H3>New Computers - Last 7 Days</H3>" | Out-String
    $disabledUsers | ConvertTo-HTML -Head $head -PreContent "<H3>Disabled Users - Last 7 Days</H3>" | Out-String
    #$groupChanges | ConvertTo-HTML -Head $head -PreContent "<H3>Group Membership Changes - Last 7 Days</H3>" | Out-String
)

# Combine All HTML Sections
$body = $htmlSections -join ""

# Save HTML Report for Logs
$logdate = Get-Date -Format yyyyMMdd
$logfileHTML = "F:\Scripts\Reporting\logs\WeeklyADReports\WeeklyADReports - $logdate.html"
$body | Out-File -FilePath $logfileHTML -Encoding utf8

# Email Settings
$emailFrom = "reporting-techops@corp.example.com"
$emailTo = "reporting-techops@corp.example.com"
$subject = "Weekly AD Report - $(Get-Date -Format 'MM/dd/yyyy')"
$smtpserver = "smtp-relay.gmail.com"

# Send Email
Send-MailMessage -To $emailTo -From $emailFrom -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver