Write-Output "Starting Program Transcript...`n"
$TranscriptStartDate = Get-Date -f 'yyyyMMdd'
$TranscriptPath = "C:\ServerLogs\EventCollectionShell-LogTranscript\${TranscriptStartDate}-EventCollectionLogTranscript.txt"
Start-Transcript -Path $TranscriptPath -NoClobber -IncludeInvocationHeader
Write-Output ""
Write-Output "GODLIKE SYSTEMS`n"
Write-Output "INITIALIZING PROGRAM...`n"

Write-Output "Welcome to the Weekly Event Log Collection Shell!`n"

Write-Output "Importing PowerShell Module to Import/Export Excel Spreadsheets...`n"
Write-Output "If You Do Not Have ImportExcel Installed Please See Below:`n"
Write-Output "ImportExcel: https://www.powershellgallery.com/packages/ImportExcel/7.8.6`n"
Write-Output "Install-Module -Name ImportExcel -RequiredVersion 7.8.6"

Import-Module ImportExcel
Write-Output ""
Write-Output "ImportExcel Successfully Imported...`n"

Write-Output "Starting Log Collection For Previous 7 Days As Of $(Get-Date)`n"

# DEFINE SERVER NAMES & DESCRIPTIONS
$servers = @(
    @{ Name = 'GODLIKE'; Description = 'SYSTEMS'; },
    @{ Name = 'CYBER'; Description = 'CONFIDENTIALITY, INTEGRITY, AVAILABILITY'; },
    @{ Name = 'UNIVERSITY OF KENTUCKY'; Description = 'GO BIG BLUE!'; }
)

# MAIN FUNCTION 0: OUTPUT EVENT LOGS IN CSV FORMAT

# DEFINE LOG TYPES & FILTERS
$logTypes = @(
    @{ LogFile = 'Application'; Types = @('Information', 'Warning', 'Error') },
    @{ LogFile = 'Security'; Types = @('Audit Failure') },
    @{ LogFile = 'System'; Types = @('Information', 'Warning', 'Error') }
)

# SET $StartDate VARIABLE TO 7 DAYS BEFORE TODAY
$StartDate = (Get-Date).AddDays(-7)

foreach ($server in $servers) {
    try {
        Write-Output "Retrieving Logs For $($server.Name) - $($server.Description)`n"
        
        $results = @{ }

        # RETRIEVE & FILTER VARIOUS EVENT LOGS FROM ARRAY OF CURRENT SERVERS
        foreach ($logType in $logTypes) {
            $results[$logType.LogFile] = Get-WmiObject -Class Win32_NTLogEvent -ComputerName $server.Name -Filter "LogFile = '$($logType.LogFile)'" | 
            Where-Object { $_.Type -in $logType.Types -and [datetime]::ParseExact($_.TimeGenerated, "yyyyMMddHHmmss.ffffff-000", $null) -ge $StartDate }
        }

        # DISPLAY LOG RESULTS
        Write-Output "Displaying Logs For $($server.Name)...`n" 
        foreach ($logType in $logTypes) {
            $results[$logType.LogFile]
        }

        # CREATE VARIABLE FOR CURRENT DATE LOGS WERE GATHERED & APPEND
        $ExportTimeStringConversion = (Get-Date).ToString("yyyy-MM-dd")
        
        foreach ($logType in $logTypes) {
            $filePath = "C:\ServerLogs\CSV\$ExportTimeStringConversion" + "_$($server.Name)$($logType.LogFile)Logs.csv"
            
            # EXPORT CSV FILE TO SPECIFIED PATH WITH UTF-8 ENCODING
            $results[$logType.LogFile] | Select-Object ComputerName, Logfile, Type, EventCode, SourceName, Message, TimeWritten | Export-CSV -Path $filePath -NoTypeInformation -Encoding UTF8
            
            Write-Output "Exported $($logType.LogFile) Logs CSV File Format: $filePath`n"
        }

        Write-Output "$($server.Name) Completed:" (Get-Date)
    } catch {
        Write-Output "Error handling logs for $($server.Name): $_.Exception.Message"
    }
}

# MAIN FUNCTION 1: APPLY PIVOT TABLES TO XLSX FORMAT

try {
    Write-Output "Applying Pivot Tables to XLSX Format`n"

    foreach ($server in $servers) {
        foreach ($logType in $logTypes) {
            # Define the CSV file path
            $CSVFilePath = "C:\ServerLogs\CSV\$ExportTimeStringConversion`_$($server.Name)$($logType.LogFile)Logs.csv"
            
            # CHECK CSV FILE EXISTENCE
            if (Test-Path -Path $CSVFilePath) {
                # DEFINE OUTPUT OF XLSX FILE PATH
                $XLSXFileName = "$($ExportTimeStringConversion)_$($server.Name)_$($logType.LogFile)Logs.xlsx"
                $XLSXFilePath = "C:\ServerLogs\XLSX\$XLSXFileName"

                # CONVERT CSV TO XLSX & APPLY PIVOT TABLE
                Import-Csv -Path $CSVFilePath | 
                Export-Excel -Path $XLSXFilePath -AutoSize -AutoFilter -IncludePivotTable -PivotRows EventCode, SourceName, Message, TimeWritten -PivotData @{'EventCode'='Count'}
                
                Write-Output "Created Pivot Table for $($logType.LogFile) Logs XLSX Format: $XLSXFilePath`n"
            } else {
                Write-Output "CSV File Not Found: $CSVFilePath`n"
            }
        }
    }
} catch {
    Write-Output "Error applying pivot tables: $_.Exception.Message"
}

# MAIN FUNCTION 2: CREATE SUMMARY TXT FILE FOR ALL SERVERS IN ONE DOCUMENT

try {
    Write-Output "Creating Summary Text File For All Servers...`n"

    $SummaryOutput = @()

    foreach ($server in $servers) {
        foreach ($logType in $logTypes) {
            $CSVFilePath = "C:\ServerLogs\CSV\$ExportTimeStringConversion`_$($server.Name)$($logType.LogFile)Logs.csv"
            
            $logSummary = Import-Csv -Path $CSVFilePath | Group-Object -Property EventCode | Select-Object Name, @{Name='Count'; Expression={$_.Count}}, @{Name='Messages'; Expression={($_.Group | Select-Object -ExpandProperty Message) -join '; '}}
            
            $summaryOutput += "`n$($server.Name) - $($logType.LogFile) Summary:`n" + ($logSummary | Out-String)
        }
    }

    # DEFINE SUMMARY OUTPUT PATH
    $SummaryFilePath = "C:\ServerLogs\SUMMARY\WeeklyEventLogSummary_$ExportTimeStringConversion.txt"

    # EXPORT SUMMARY TO TEXT FILE
    $SummaryOutput | Out-File -FilePath $SummaryFilePath -Encoding UTF8
    Write-Output "Summary Exported To: $SummaryFilePath`n"
} catch {
    Write-Output "Error creating summary text file: $_.Exception.Message"
}

Write-Output "All Operations Completed Successfully!`n"

# END OF PROGRAM | EXPORT PROGRAM TRANSCRIPT
Write-Output "Program Transcript Writing To... $TranscriptPath`n"
Stop-Transcript
