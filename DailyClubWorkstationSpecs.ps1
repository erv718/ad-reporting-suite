$masterPath = '\\SYSADMIN01\Reports\MasterClubWorkstationSpecs.csv'
$dailyPath = '\\SYSADMIN01\Reports\DailyClubWorkstationSpecs.csv'

# Import the data
$masterData = Import-Csv -Path $masterPath
$dailyData = Import-Csv -Path $dailyPath

$today = Get-Date -Format "yyyy-MM-dd"

# Iterate over each unique hostname in master data
foreach ($hostname in ($masterData.Hostname | Sort-Object | Get-Unique)) {
    # Get the latest data for this hostname from master data
    $latestEntry = $masterData | Where-Object { $_.Hostname -eq $hostname } | Sort-Object Date -Descending | Select-Object -First 1

    # Check if this hostname exists in daily data
    $existingDailyEntry = $dailyData | Where-Object { $_.Hostname -eq $hostname }

    if ($existingDailyEntry) {
        # If it exists, update the entry with the latest data
        $dailyData = $dailyData | Where-Object { $_.Hostname -ne $hostname }
        $dailyData += $latestEntry
    } else {
        # If not, just append the latest data
        $dailyData += $latestEntry
    }
}

# Write back the refreshed data to DailyClubWorkstationSpecs.csv
$dailyData | Export-Csv -Path $dailyPath -NoTypeInformation -Force
