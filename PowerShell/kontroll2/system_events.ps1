$startDate = Read-Host "Enter start date (e.g., 'MM/DD/YYYY')"
$endDate = Read-Host "Enter end date (e.g., 'MM/DD/YYYY')"
$filePath = Read-Host "Enter path to save the file (e.g., 'C:\path\to\file.txt')"

$startTime = [datetime]::ParseExact($startDate, 'MM/dd/yyyy', $null)
$endTime = [datetime]::ParseExact($endDate, 'MM/dd/yyyy', $null)

$logs = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    StartTime = $startTime
    EndTime = $endTime
    Level = @(1, 2, 3) 
} -ErrorAction SilentlyContinue

$groupedLogs = $logs | Group-Object { $_.ProviderName } | 
               Sort-Object { $_.Count } -Descending | 
               ForEach-Object {
                   $group = $_
                   $events = $group.Group | 
                             Sort-Object TimeCreated -Descending | 
                             Select-Object @{Name="Time"; Expression={$_.TimeCreated}}, 
                                           @{Name="Message"; Expression={$_.Message.Substring(0, [Math]::Min(100, $_.Message.Length))}}

                   "[ Title: $($group.Name) ]`n" + 
                   ($events | Format-Table -HideTableHeaders | Out-String).Trim()
               }

If (!(Test-Path -Path $(Split-Path -Path $filePath -Parent))) {
    New-Item -ItemType Directory -Force -Path $(Split-Path -Path $filePath -Parent)
}

$groupedLogs | Out-File -FilePath $filePath

Write-Host "Logs saved to: $filePath"
