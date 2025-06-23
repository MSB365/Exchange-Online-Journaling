<#
.SYNOPSIS
    Configure Exchange Online Journaling and Generate Monthly Reports
.DESCRIPTION
    This script configures journaling for all incoming and outgoing messages in Exchange Online
    and generates monthly HTML reports of message activity.
.PARAMETER JournalEmailAddress
    Email address where journal reports will be sent
.PARAMETER UndeliverableReportsAddress
    Email address where undeliverable journal reports will be sent (required for journaling)
.PARAMETER ReportPath
    Path where HTML reports will be saved
.PARAMETER MonthsBack
    Number of months back to generate reports for (default: 1)
.PARAMETER UseHistoricalSearch
    Use historical search for data older than 10 days (slower but more comprehensive)
.PARAMETER SkipJournalingConfig
    Skip journaling configuration and only generate reports
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$JournalEmailAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$UndeliverableReportsAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "C:\MDM\journaling\ExchangeReports",
    
    [Parameter(Mandatory=$false)]
    [int]$MonthsBack = 1,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseHistoricalSearch = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipJournalingConfig = $false
)

# Import required modules
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host "✓ Exchange Online Management module imported successfully" -ForegroundColor Green
} catch {
    Write-Error "Failed to import Exchange Online Management module. Please install it using: Install-Module -Name ExchangeOnlineManagement"
    exit 1
}

# Function to connect to Exchange Online
function Connect-ExchangeOnlineSecure {
    try {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
        Connect-ExchangeOnline -ShowProgress $true
        Write-Host "✓ Connected to Exchange Online successfully" -ForegroundColor Green
        return $true
    } catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        return $false
    }
}

# Function to validate mailbox exists
function Test-MailboxExists {
    param([string]$EmailAddress)
    
    try {
        $mailbox = Get-Mailbox -Identity $EmailAddress -ErrorAction SilentlyContinue
        return ($mailbox -ne $null)
    } catch {
        return $false
    }
}

# Function to configure undeliverable journal reports address
function Set-UndeliverableJournalReportsAddress {
    param([string]$UndeliverableAddress)
    
    Write-Host "Configuring undeliverable journal reports address..." -ForegroundColor Yellow
    
    try {
        # Validate the undeliverable address mailbox exists
        if (-not (Test-MailboxExists -EmailAddress $UndeliverableAddress)) {
            Write-Error "Undeliverable reports mailbox '$UndeliverableAddress' not found."
            Write-Host "Please ensure this mailbox exists before configuring journaling." -ForegroundColor Red
            return $false
        }
        
        # Set the undeliverable journal reports address
        Write-Host "Setting undeliverable journal reports address to: $UndeliverableAddress" -ForegroundColor Cyan
        Set-TransportConfig -JournalingReportNdrTo $UndeliverableAddress
        
        # Wait a moment for the setting to propagate
        Start-Sleep -Seconds 5
        
        # Verify the setting
        $TransportConfig = Get-TransportConfig
        if ($TransportConfig.JournalingReportNdrTo -eq $UndeliverableAddress) {
            Write-Host "✓ Undeliverable journal reports address configured successfully" -ForegroundColor Green
            return $true
        } else {
            Write-Error "Failed to verify undeliverable journal reports address configuration"
            return $false
        }
        
    } catch {
        Write-Error "Failed to configure undeliverable journal reports address: $($_.Exception.Message)"
        return $false
    }
}

# Function to configure journaling with proper prerequisites
function Configure-Journaling {
    param(
        [string]$JournalEmail,
        [string]$UndeliverableAddress
    )
    
    Write-Host "`nConfiguring Exchange Online Journaling..." -ForegroundColor Yellow
    
    try {
        # Step 1: Validate journal mailbox exists
        Write-Host "Validating journal mailbox..." -ForegroundColor Cyan
        if (-not (Test-MailboxExists -EmailAddress $JournalEmail)) {
            Write-Error "Journal mailbox '$JournalEmail' not found."
            Write-Host "Please create this mailbox before configuring journaling." -ForegroundColor Red
            return $false
        }
        Write-Host "✓ Journal mailbox validated: $JournalEmail" -ForegroundColor Green
        
        # Step 2: Check current undeliverable reports configuration
        Write-Host "Checking current transport configuration..." -ForegroundColor Cyan
        $TransportConfig = Get-TransportConfig
        $CurrentUndeliverableAddress = $TransportConfig.JournalingReportNdrTo
        
        if ($CurrentUndeliverableAddress) {
            Write-Host "✓ Undeliverable journal reports address already configured: $CurrentUndeliverableAddress" -ForegroundColor Green
        } else {
            Write-Host "⚠ No undeliverable journal reports address configured" -ForegroundColor Yellow
            
            # Step 3: Configure undeliverable address if provided
            if ($UndeliverableAddress) {
                if (-not (Set-UndeliverableJournalReportsAddress -UndeliverableAddress $UndeliverableAddress)) {
                    Write-Error "Failed to configure undeliverable reports address. Cannot proceed with journal rule creation."
                    return $false
                }
            } else {
                # Prompt user for undeliverable address
                Write-Host "`nJournaling requires an undeliverable reports address to be configured." -ForegroundColor Yellow
                Write-Host "This mailbox will receive journal reports that cannot be delivered to the main journal mailbox." -ForegroundColor Cyan
                
                do {
                    $UserInput = Read-Host "Enter an email address for undeliverable journal reports (or 'skip' to skip journaling configuration)"
                    
                    if ($UserInput -eq 'skip') {
                        Write-Host "Skipping journaling configuration as requested." -ForegroundColor Yellow
                        return $false
                    }
                    
                    if ($UserInput) {
                        if (Test-MailboxExists -EmailAddress $UserInput) {
                            if (Set-UndeliverableJournalReportsAddress -UndeliverableAddress $UserInput) {
                                break
                            }
                        } else {
                            Write-Host "Mailbox '$UserInput' not found. Please enter a valid mailbox address." -ForegroundColor Red
                        }
                    }
                } while ($true)
            }
        }
        
        # Step 4: Create or update journaling rule
        $ruleName = "All-Messages-Journal-Rule"
        Write-Host "Configuring journal rule: $ruleName" -ForegroundColor Cyan
        
        # Check if rule already exists
        $existingRule = Get-JournalRule -Identity $ruleName -ErrorAction SilentlyContinue
        
        if ($existingRule) {
            Write-Host "Journal rule '$ruleName' already exists. Updating..." -ForegroundColor Yellow
            try {
                Set-JournalRule -Identity $ruleName -JournalEmailAddress $JournalEmail -Scope Global -Enabled $true -Confirm:$false
                Write-Host "✓ Journal rule updated successfully" -ForegroundColor Green
            } catch {
                Write-Error "Failed to update journal rule: $($_.Exception.Message)"
                return $false
            }
        } else {
            Write-Host "Creating new journal rule '$ruleName'..." -ForegroundColor Yellow
            try {
                New-JournalRule -Name $ruleName -JournalEmailAddress $JournalEmail -Scope Global -Enabled $true -Confirm:$false
                Write-Host "✓ Journal rule created successfully" -ForegroundColor Green
            } catch {
                Write-Error "Failed to create journal rule: $($_.Exception.Message)"
                Write-Host "This might be due to insufficient permissions or missing prerequisites." -ForegroundColor Red
                return $false
            }
        }
        
        # Step 5: Verify the rule configuration
        Start-Sleep -Seconds 3
        $rule = Get-JournalRule -Identity $ruleName -ErrorAction SilentlyContinue
        if ($rule) {
            Write-Host "`n✓ Journaling configuration completed successfully:" -ForegroundColor Green
            Write-Host "  - Rule Name: $($rule.Name)" -ForegroundColor Cyan
            Write-Host "  - Journal Email: $($rule.JournalEmailAddress)" -ForegroundColor Cyan
            Write-Host "  - Scope: $($rule.Scope)" -ForegroundColor Cyan
            Write-Host "  - Enabled: $($rule.Enabled)" -ForegroundColor Cyan
            
            # Show current transport config
            $FinalTransportConfig = Get-TransportConfig
            Write-Host "  - Undeliverable Reports: $($FinalTransportConfig.JournalingReportNdrTo)" -ForegroundColor Cyan
            
            return $true
        } else {
            Write-Error "Failed to verify journal rule creation"
            return $false
        }
        
    } catch {
        Write-Error "Failed to configure journaling: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception)" -ForegroundColor Red
        return $false
    }
}

# Function to get message trace data for full month
function Get-MessageTraceData {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [bool]$UseHistoricalSearch = $false
    )
    
    $Today = Get-Date
    $TenDaysAgo = $Today.AddDays(-10).Date
    
    Write-Host "Retrieving message trace data from $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))..." -ForegroundColor Yellow
    
    # Check if we need historical search
    $NeedsHistoricalSearch = $StartDate -lt $TenDaysAgo
    
    if ($NeedsHistoricalSearch -and -not $UseHistoricalSearch) {
        Write-Warning "The requested date range requires historical data (older than 10 days)."
        Write-Host "For complete monthly reports, use -UseHistoricalSearch parameter." -ForegroundColor Yellow
        Write-Host "Proceeding with available recent data only..." -ForegroundColor Cyan
        
        # Adjust to available data range but keep it meaningful
        $AdjustedStartDate = [math]::Max($StartDate.Ticks, $TenDaysAgo.Ticks) | ForEach-Object { New-Object DateTime $_ }
        $AdjustedEndDate = [math]::Min($EndDate.Ticks, $Today.AddDays(-1).Ticks) | ForEach-Object { New-Object DateTime $_ }
        
        Write-Host "Adjusted range: $($AdjustedStartDate.ToString('yyyy-MM-dd')) to $($AdjustedEndDate.ToString('yyyy-MM-dd'))" -ForegroundColor Green
        $StartDate = $AdjustedStartDate
        $EndDate = $AdjustedEndDate
    }
    
    $allMessages = @()
    
    try {
        if ($UseHistoricalSearch -and $NeedsHistoricalSearch) {
            Write-Host "Using historical search for complete monthly data..." -ForegroundColor Yellow
            $allMessages = Get-HistoricalMessageTrace -StartDate $StartDate -EndDate $EndDate
        } else {
            # Use regular message trace
            $pageSize = 5000
            $page = 1
            
            Write-Host "Retrieving messages using Get-MessageTrace..." -ForegroundColor Cyan
            
            do {
                try {
                    Write-Host "  Fetching page $page..." -ForegroundColor Gray
                    $messages = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -PageSize $pageSize -Page $page
                    if ($messages) {
                        $allMessages += $messages
                        Write-Host "  Retrieved $($messages.Count) messages (Page $page)" -ForegroundColor Cyan
                        $page++
                        
                        # Add delay to avoid throttling
                        if ($messages.Count -eq $pageSize) {
                            Start-Sleep -Seconds 2
                        }
                    }
                } catch {
                    Write-Warning "Error retrieving messages on page: $($_.Exception.Message)"
                    break
                }
            } while ($messages.Count -eq $pageSize)
        }
        
        Write-Host "✓ Total messages retrieved: $($allMessages.Count)" -ForegroundColor Green
        return $allMessages
        
    } catch {
        Write-Error "Failed to retrieve message trace data: $($_.Exception.Message)"
        return @()
    }
}

# Function to get historical message trace data
function Get-HistoricalMessageTrace {
    param(
        [DateTime]$StartDate,
        [DateTime]$EndDate
    )
    
    Write-Host "Starting historical message trace search..." -ForegroundColor Yellow
    Write-Host "Note: Historical searches can take several minutes to complete." -ForegroundColor Cyan
    
    try {
        # Start a historical search
        $SearchName = "ExchangeJournalingReport-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
        
        $HistoricalSearch = Start-HistoricalSearch -ReportTitle $SearchName -StartDate $StartDate -EndDate $EndDate -ReportType MessageTrace
        
        if ($HistoricalSearch) {
            Write-Host "Historical search started: $($HistoricalSearch.JobId)" -ForegroundColor Green
            Write-Host "Waiting for search to complete..." -ForegroundColor Yellow
            
            # Wait for the search to complete
            do {
                Start-Sleep -Seconds 30
                $SearchStatus = Get-HistoricalSearch -JobId $HistoricalSearch.JobId
                Write-Host "Search status: $($SearchStatus.Status)" -ForegroundColor Cyan
            } while ($SearchStatus.Status -eq "InProgress")
            
            if ($SearchStatus.Status -eq "Done") {
                Write-Host "✓ Historical search completed successfully" -ForegroundColor Green
                
                # Get the results
                $Results = Get-MessageTrace -MessageTraceId $HistoricalSearch.JobId
                return $Results
            } else {
                Write-Warning "Historical search failed with status: $($SearchStatus.Status)"
                return @()
            }
        } else {
            Write-Error "Failed to start historical search"
            return @()
        }
        
    } catch {
        Write-Error "Historical search failed: $($_.Exception.Message)"
        Write-Host "Falling back to recent data only..." -ForegroundColor Yellow
        
        # Fallback to recent data
        $RecentStartDate = (Get-Date).AddDays(-9)
        return Get-MessageTrace -StartDate $RecentStartDate -EndDate $EndDate
    }
}

# Function to get detailed user information
function Get-UserDetails {
    param([string]$EmailAddress)
    
    try {
        $user = Get-User -Identity $EmailAddress -ErrorAction SilentlyContinue
        if ($user) {
            return @{
                DisplayName = $user.DisplayName
                Department = $user.Department
                Office = $user.Office
                Title = $user.Title
                Company = $user.Company
                City = $user.City
                Country = $user.CountryOrRegion
            }
        }
    } catch {
        # Silently handle errors for external users
    }
    
    return @{
        DisplayName = $EmailAddress
        Department = "External/Unknown"
        Office = ""
        Title = ""
        Company = ""
        City = ""
        Country = ""
    }
}

# Function to generate enhanced HTML report
function Generate-HTMLReport {
    param(
        [array]$Messages,
        [DateTime]$StartDate,
        [DateTime]$EndDate,
        [string]$OutputPath
    )
    
    Write-Host "Generating enhanced HTML report..." -ForegroundColor Yellow
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    # Handle empty message array
    if ($Messages.Count -eq 0) {
        Write-Warning "No messages found for the specified date range."
        
        $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Online Monthly Report - $($StartDate.ToString('MMMM yyyy'))</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 800px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; text-align: center; }
        .no-data { text-align: center; color: #666; font-size: 1.2em; margin: 40px 0; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📧 Exchange Online Monthly Report</h1>
        <p style="text-align: center; color: #666;">
            Report Period: <strong>$($StartDate.ToString('MMMM dd, yyyy')) - $($EndDate.ToString('MMMM dd, yyyy'))</strong>
        </p>
        <div class="no-data">
            <h2>No Data Available</h2>
            <p>No messages were found for the specified date range.</p>
        </div>
        <div style="text-align: center; margin-top: 30px; color: #666; font-size: 0.9em;">
            <p>Report generated on $(Get-Date -Format 'MMMM dd, yyyy HH:mm:ss')</p>
        </div>
    </div>
</body>
</html>
"@
        
        $fileName = "ExchangeOnline-Report-$($StartDate.ToString('yyyy-MM'))-NoData.html"
        $filePath = Join-Path $OutputPath $fileName
        $html | Out-File -FilePath $filePath -Encoding UTF8
        
        Write-Host "✓ No-data report saved to: $filePath" -ForegroundColor Yellow
        return $filePath
    }
    
    # Calculate statistics
    Write-Host "Calculating message statistics..." -ForegroundColor Gray
    $totalMessages = $Messages.Count
    $incomingMessages = ($Messages | Where-Object { $_.Direction -eq "Inbound" }).Count
    $outgoingMessages = ($Messages | Where-Object { $_.Direction -eq "Outbound" }).Count
    
    # Group by status
    $statusGroups = $Messages | Where-Object { $_.Status } | Group-Object Status
    
    # Group by date for daily statistics
    Write-Host "Processing daily statistics..." -ForegroundColor Gray
    $dailyStats = $Messages | Where-Object { $_.Received } | Group-Object { 
        $_.Received.Date.ToString('yyyy-MM-dd')
    } | Sort-Object Name
    
    # Get detailed top senders with user information
    Write-Host "Analyzing top senders..." -ForegroundColor Gray
    $topSendersRaw = $Messages | Where-Object { $_.SenderAddress } | Group-Object SenderAddress | Sort-Object Count -Descending | Select-Object -First 10
    $topSenders = @()
    
    foreach ($sender in $topSendersRaw) {
        Write-Host "  Getting details for: $($sender.Name)" -ForegroundColor DarkGray
        $userDetails = Get-UserDetails -EmailAddress $sender.Name
        $senderMessages = $Messages | Where-Object { $_.SenderAddress -eq $sender.Name }
        
        $topSenders += [PSCustomObject]@{
            EmailAddress = $sender.Name
            DisplayName = $userDetails.DisplayName
            Department = $userDetails.Department
            Title = $userDetails.Title
            Office = $userDetails.Office
            Company = $userDetails.Company
            MessageCount = $sender.Count
            InboundCount = ($senderMessages | Where-Object { $_.Direction -eq "Inbound" }).Count
            OutboundCount = ($senderMessages | Where-Object { $_.Direction -eq "Outbound" }).Count
            UniqueRecipients = ($senderMessages | Select-Object RecipientAddress -Unique).Count
            AvgPerDay = [math]::Round($sender.Count / [math]::Max(($EndDate - $StartDate).Days, 1), 1)
        }
    }
    
    # Get detailed top recipients with user information
    Write-Host "Analyzing top recipients..." -ForegroundColor Gray
    $topRecipientsRaw = $Messages | Where-Object { $_.RecipientAddress } | Group-Object RecipientAddress | Sort-Object Count -Descending | Select-Object -First 10
    $topRecipients = @()
    
    foreach ($recipient in $topRecipientsRaw) {
        Write-Host "  Getting details for: $($recipient.Name)" -ForegroundColor DarkGray
        $userDetails = Get-UserDetails -EmailAddress $recipient.Name
        $recipientMessages = $Messages | Where-Object { $_.RecipientAddress -eq $recipient.Name }
        
        $topRecipients += [PSCustomObject]@{
            EmailAddress = $recipient.Name
            DisplayName = $userDetails.DisplayName
            Department = $userDetails.Department
            Title = $userDetails.Title
            Office = $userDetails.Office
            Company = $userDetails.Company
            MessageCount = $recipient.Count
            InboundCount = ($recipientMessages | Where-Object { $_.Direction -eq "Inbound" }).Count
            OutboundCount = ($recipientMessages | Where-Object { $_.Direction -eq "Outbound" }).Count
            UniqueSenders = ($recipientMessages | Select-Object SenderAddress -Unique).Count
            AvgPerDay = [math]::Round($recipient.Count / [math]::Max(($EndDate - $StartDate).Days, 1), 1)
        }
    }
    
    # Calculate additional statistics
    $avgMessagesPerDay = [math]::Round($totalMessages / [math]::Max(($EndDate - $StartDate).Days, 1), 0)
    $uniqueSenders = ($Messages | Select-Object SenderAddress -Unique).Count
    $uniqueRecipients = ($Messages | Select-Object RecipientAddress -Unique).Count
    
    # Generate enhanced HTML content
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Online Monthly Report - $($StartDate.ToString('MMMM yyyy'))</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; border-radius: 12px; box-shadow: 0 8px 32px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 30px; text-align: center; }
        .header h1 { margin: 0; font-size: 2.5em; font-weight: 300; }
        .header p { margin: 10px 0 0 0; font-size: 1.2em; opacity: 0.9; }
        .content { padding: 30px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 30px 0; }
        .stat-card { background: linear-gradient(135deg, #667eea, #764ba2); color: white; padding: 25px; border-radius: 10px; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.1); }
        .stat-number { font-size: 2.5em; font-weight: bold; margin-bottom: 5px; }
        .stat-label { font-size: 1em; opacity: 0.9; }
        h2 { color: #333; border-bottom: 3px solid #0078d4; padding-bottom: 10px; margin-top: 40px; font-size: 1.8em; }
        .table-container { overflow-x: auto; margin: 20px 0; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        table { width: 100%; border-collapse: collapse; background: white; }
        th { background: linear-gradient(135deg, #0078d4, #106ebe); color: white; padding: 15px 12px; text-align: left; font-weight: 600; }
        td { padding: 12px; border-bottom: 1px solid #eee; }
        tr:hover { background-color: #f8f9ff; }
        .user-info { display: flex; flex-direction: column; }
        .user-name { font-weight: bold; color: #333; margin-bottom: 2px; }
        .user-email { color: #666; font-size: 0.9em; margin-bottom: 2px; }
        .user-details { color: #888; font-size: 0.8em; }
        .message-stats { text-align: center; }
        .message-count { font-weight: bold; font-size: 1.1em; color: #0078d4; }
        .message-breakdown { font-size: 0.8em; color: #666; margin-top: 2px; }
        .chart-container { margin: 20px 0; }
        .bar { height: 25px; background: linear-gradient(90deg, #0078d4, #40e0d0); margin: 3px 0; border-radius: 4px; position: relative; }
        .bar-label { position: absolute; left: 10px; top: 50%; transform: translateY(-50%); color: white; font-weight: bold; font-size: 0.9em; }
        .footer { text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; color: #666; border-radius: 8px; }
        .external-indicator { background-color: #ffc107; color: #856404; padding: 2px 6px; border-radius: 3px; font-size: 0.7em; margin-left: 5px; }
        .internal-indicator { background-color: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 0.7em; margin-left: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📧 Exchange Online Monthly Report</h1>
            <p>Report Period: $($StartDate.ToString('MMMM dd, yyyy')) - $($EndDate.ToString('MMMM dd, yyyy'))</p>
        </div>
        
        <div class="content">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-number">$totalMessages</div>
                    <div class="stat-label">Total Messages</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$incomingMessages</div>
                    <div class="stat-label">Incoming Messages</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$outgoingMessages</div>
                    <div class="stat-label">Outgoing Messages</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$avgMessagesPerDay</div>
                    <div class="stat-label">Avg Messages/Day</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$uniqueSenders</div>
                    <div class="stat-label">Unique Senders</div>
                </div>
                <div class="stat-card">
                    <div class="stat-number">$uniqueRecipients</div>
                    <div class="stat-label">Unique Recipients</div>
                </div>
            </div>
"@

    # Add message status distribution
    if ($statusGroups.Count -gt 0) {
        $html += @"
            <h2>📊 Message Status Distribution</h2>
            <div class="table-container">
                <table>
                    <tr><th>Status</th><th>Count</th><th>Percentage</th></tr>
"@
        
        foreach ($status in $statusGroups) {
            $percentage = [math]::Round(($status.Count / $totalMessages) * 100, 1)
            $statusName = if ($status.Name) { $status.Name } else { "Unknown" }
            $html += "<tr><td>$statusName</td><td>$($status.Count)</td><td>$percentage%</td></tr>"
        }
        
        $html += "</table></div>"
    }

    # Add daily message volume
    if ($dailyStats.Count -gt 0) {
        $html += @"
            <h2>📈 Daily Message Volume</h2>
            <div class="chart-container">
                <div class="table-container">
                    <table>
                        <tr><th>Date</th><th>Messages</th><th>Visual Distribution</th></tr>
"@
        
        $maxDaily = ($dailyStats | Measure-Object Count -Maximum).Maximum
        
        foreach ($day in $dailyStats) {
            $date = [DateTime]::Parse($day.Name)
            $displayDate = $date.ToString('MMM dd, yyyy')
            $barWidth = if ($maxDaily -gt 0) { [math]::Round(($day.Count / $maxDaily) * 100, 1) } else { 0 }
            $html += "<tr><td>$displayDate</td><td class='message-count'>$($day.Count)</td><td><div class='bar' style='width: $barWidth%;'><span class='bar-label'>$($day.Count)</span></div></td></tr>"
        }
        
        $html += "</table></div></div>"
    }

    # Add detailed top senders
    $html += @"
        <h2>👤 Top 10 Senders - Detailed Analysis</h2>
        <div class="table-container">
            <table>
                <tr>
                    <th>User Information</th>
                    <th>Message Statistics</th>
                    <th>Activity Details</th>
                    <th>Daily Average</th>
                </tr>
"@

    foreach ($sender in $topSenders) {
        $isExternal = $sender.Department -eq "External/Unknown"
        $indicator = if ($isExternal) { "<span class='external-indicator'>EXTERNAL</span>" } else { "<span class='internal-indicator'>INTERNAL</span>" }
        
        $userInfo = @"
<div class="user-info">
    <div class="user-name">$($sender.DisplayName)$indicator</div>
    <div class="user-email">$($sender.EmailAddress)</div>
    <div class="user-details">
"@
        
        if (-not $isExternal) {
            if ($sender.Title) { $userInfo += "$($sender.Title)<br>" }
            if ($sender.Department) { $userInfo += "$($sender.Department)<br>" }
            if ($sender.Office) { $userInfo += "$($sender.Office)" }
        }
        
        $userInfo += "</div></div>"
        
        $messageStats = @"
<div class="message-stats">
    <div class="message-count">$($sender.MessageCount)</div>
    <div class="message-breakdown">
        In: $($sender.InboundCount) | Out: $($sender.OutboundCount)
    </div>
</div>
"@
        
        $activityDetails = "Unique Recipients: $($sender.UniqueRecipients)"
        $dailyAvg = "$($sender.AvgPerDay) msg/day"
        
        $html += "<tr><td>$userInfo</td><td>$messageStats</td><td>$activityDetails</td><td>$dailyAvg</td></tr>"
    }

    $html += "</table></div>"

    # Add detailed top recipients
    $html += @"
        <h2>📬 Top 10 Recipients - Detailed Analysis</h2>
        <div class="table-container">
            <table>
                <tr>
                    <th>User Information</th>
                    <th>Message Statistics</th>
                    <th>Activity Details</th>
                    <th>Daily Average</th>
                </tr>
"@

    foreach ($recipient in $topRecipients) {
        $isExternal = $recipient.Department -eq "External/Unknown"
        $indicator = if ($isExternal) { "<span class='external-indicator'>EXTERNAL</span>" } else { "<span class='internal-indicator'>INTERNAL</span>" }
        
        $userInfo = @"
<div class="user-info">
    <div class="user-name">$($recipient.DisplayName)$indicator</div>
    <div class="user-email">$($recipient.EmailAddress)</div>
    <div class="user-details">
"@
        
        if (-not $isExternal) {
            if ($recipient.Title) { $userInfo += "$($recipient.Title)<br>" }
            if ($recipient.Department) { $userInfo += "$($recipient.Department)<br>" }
            if ($recipient.Office) { $userInfo += "$($recipient.Office)" }
        }
        
        $userInfo += "</div></div>"
        
        $messageStats = @"
<div class="message-stats">
    <div class="message-count">$($recipient.MessageCount)</div>
    <div class="message-breakdown">
        In: $($recipient.InboundCount) | Out: $($recipient.OutboundCount)
    </div>
</div>
"@
        
        $activityDetails = "Unique Senders: $($recipient.UniqueSenders)"
        $dailyAvg = "$($recipient.AvgPerDay) msg/day"
        
        $html += "<tr><td>$userInfo</td><td>$messageStats</td><td>$activityDetails</td><td>$dailyAvg</td></tr>"
    }

    # Close the HTML
    $html += @"
            </table>
        </div>
        
        <div class="footer">
            <p><strong>Report generated on $(Get-Date -Format 'MMMM dd, yyyy HH:mm:ss')</strong></p>
            <p>Exchange Online Journaling & Reporting Script</p>
            <p>Total analysis time: $((Get-Date) - $StartDate | Select-Object -ExpandProperty TotalMinutes | ForEach-Object { [math]::Round($_, 1) }) minutes</p>
        </div>
    </div>
</body>
</html>
"@

    # Save the report
    try {
        $fileName = "ExchangeOnline-Report-$($StartDate.ToString('yyyy-MM')).html"
        $filePath = Join-Path $OutputPath $fileName
        $html | Out-File -FilePath $filePath -Encoding UTF8
        
        Write-Host "✓ Enhanced HTML report saved to: $filePath" -ForegroundColor Green
        return $filePath
    } catch {
        Write-Error "Failed to save HTML report: $($_.Exception.Message)"
        return $null
    }
}

# Main execution
function Main {
    Write-Host "=== Exchange Online Journaling Configuration & Reporting ===" -ForegroundColor Cyan
    Write-Host "Journal Email: $JournalEmailAddress" -ForegroundColor White
    if ($UndeliverableReportsAddress) {
        Write-Host "Undeliverable Reports: $UndeliverableReportsAddress" -ForegroundColor White
    }
    Write-Host "Report Path: $ReportPath" -ForegroundColor White
    Write-Host "Months Back: $MonthsBack" -ForegroundColor White
    Write-Host "Skip Journaling Config: $SkipJournalingConfig" -ForegroundColor White
    Write-Host ""
    
    # Connect to Exchange Online
    if (-not (Connect-ExchangeOnlineSecure)) {
        exit 1
    }
    
    # Configure journaling (unless skipped)
    if (-not $SkipJournalingConfig) {
        $JournalingSuccess = Configure-Journaling -JournalEmail $JournalEmailAddress -UndeliverableAddress $UndeliverableReportsAddress
        if (-not $JournalingSuccess) {
            Write-Host "`nJournaling configuration failed or was skipped." -ForegroundColor Yellow
            Write-Host "You can run the script with -SkipJournalingConfig to only generate reports." -ForegroundColor Cyan
            
            $ContinueChoice = Read-Host "Continue with report generation anyway? (y/n)"
            if ($ContinueChoice -ne 'y' -and $ContinueChoice -ne 'Y') {
                Write-Host "Script execution cancelled." -ForegroundColor Yellow
                Disconnect-ExchangeOnline -Confirm:$false
                exit 0
            }
        }
    } else {
        Write-Host "Skipping journaling configuration as requested." -ForegroundColor Yellow
    }
    
    # Generate monthly reports
    if ($MonthsBack -gt 0) {
        for ($i = 1; $i -le $MonthsBack; $i++) {
            $currentDate = Get-Date
            
            # Calculate full month dates
            if ($i -eq 1) {
                # Current month (from 1st to today or end of month if month is complete)
                $startDate = New-Object DateTime($currentDate.Year, $currentDate.Month, 1)
                $endDate = if ($currentDate.Day -eq [DateTime]::DaysInMonth($currentDate.Year, $currentDate.Month)) {
                    $currentDate.Date
                } else {
                    $startDate.AddMonths(1).AddDays(-1)
                }
            } else {
                # Previous months (full months)
                $targetMonth = $currentDate.AddMonths(-$i)
                $startDate = New-Object DateTime($targetMonth.Year, $targetMonth.Month, 1)
                $endDate = $startDate.AddMonths(1).AddDays(-1)
            }
            
            Write-Host "`n--- Generating Report for $($startDate.ToString('MMMM yyyy')) ---" -ForegroundColor Cyan
            Write-Host "Date range: $($startDate.ToString('yyyy-MM-dd')) to $($endDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
            
            # Get message trace data
            $messages = Get-MessageTraceData -StartDate $startDate -EndDate $endDate -UseHistoricalSearch $UseHistoricalSearch
            
            # Generate HTML report
            $reportPath = Generate-HTMLReport -Messages $messages -StartDate $startDate -EndDate $endDate -OutputPath $ReportPath
            if ($reportPath) {
                Write-Host "Report completed: $reportPath" -ForegroundColor Green
            }
        }
    }
    
    # Disconnect from Exchange Online
    try {
        Disconnect-ExchangeOnline -Confirm:$false
        Write-Host "`n✓ Disconnected from Exchange Online" -ForegroundColor Green
    } catch {
        Write-Warning "Error disconnecting from Exchange Online: $($_.Exception.Message)"
    }
    
    Write-Host "`n=== Script Execution Completed ===" -ForegroundColor Cyan
    Write-Host "`nReport Features:" -ForegroundColor Yellow
    Write-Host "• Full monthly date ranges" -ForegroundColor White
    Write-Host "• Detailed user information (name, title, department, office)" -ForegroundColor White
    Write-Host "• Message statistics breakdown (inbound/outbound counts)" -ForegroundColor White
    Write-Host "• Activity analysis (unique senders/recipients, daily averages)" -ForegroundColor White
    Write-Host "• External vs internal user identification" -ForegroundColor White
    Write-Host "• Enhanced visual design with charts and graphs" -ForegroundColor White
}

# Execute main function
Main
