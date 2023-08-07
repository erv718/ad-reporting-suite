# AD Reporting Suite

There was no visibility into what was changing in Active Directory week to week. New accounts, group membership changes, stale workstations, expiring contractor accounts - all of it required someone to manually check.

## What These Scripts Do

- **WeeklyADReport.ps1** - Generates an HTML report of all AD changes from the past 7 days: new users, new groups, new computers, group membership changes, deleted objects, and who created them. Sends via email.
- **EmailManagersForExpiringAccounts.ps1** - Finds enabled accounts expiring within 33 days, groups them by manager, and emails each manager a formatted list of their expiring direct reports.
- **OldADComputersReporting.ps1** - Identifies stale computer objects that haven't authenticated in a defined period.
- **DailyClubWorkstationSpecs.ps1** - Collects hardware specs from club workstations daily for inventory tracking.
- **Paylocity_Report.ps1** - Cross-references Paylocity HR data against AD to flag discrepancies.
- **get_all_ad_users.ps1** - Quick export of all AD user objects to CSV.
- **PSHTML-AD.ps1** - Full HTML dashboard of AD environment using the PSHTML module.
- **ADReport.ps1** - Lightweight AD summary report.

## Usage

```powershell
# Run the weekly report (typically scheduled as a task)
.\WeeklyADReport.ps1

# Email managers about expiring accounts
.\EmailManagersForExpiringAccounts.ps1

# Check for stale computer objects
.\OldADComputersReporting.ps1
```

Most of these are designed to run as scheduled tasks on a domain controller or management server.

## Requirements

- PowerShell 5.1+
- `ActiveDirectory` module (RSAT)
- SMTP relay access for email-sending scripts
- `PSHTML` module for the HTML dashboard script
- `ImportExcel` module for Paylocity report

## Blog Post

[Building an Automated Active Directory Reporting Pipeline](https://blog.soarsystems.cc/active-directory-reporting-suite/)
