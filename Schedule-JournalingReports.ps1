<#
.SYNOPSIS
    Schedule Exchange Online Journaling Reports
.DESCRIPTION
    Creates a scheduled task to run monthly journaling reports automatically
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$JournalEmailAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$ReportPath = "C:\MDM\journaling\ExchangeReports",
    
    [Parameter(Mandatory=$false)]
    [string]$ScriptPath = "C:\MDM\journaling\Configure-ExchangeJournaling.ps1",
    
    [Parameter(Mandatory=$false)]
    [string]$TaskName = "Exchange Online Monthly Report"
)

# Function to create scheduled task
function Create-ScheduledTask {
    param(
        [string]$TaskName,
        [string]$ScriptPath,
        [string]$JournalEmail,
        [string]$ReportPath
    )
    
    try {
        Write-Host "Creating scheduled task '$TaskName'..." -ForegroundColor Yellow
        
        # Check if task already exists
        $ExistingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($ExistingTask) {
            Write-Host "Task '$TaskName' already exists. Removing old task..." -ForegroundColor Yellow
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        }
        
        # Define the action
        $ActionArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -JournalEmailAddress `"$JournalEmail`" -ReportPath `"$ReportPath`""
        $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $ActionArgs
        
        # Create monthly trigger using XML definition (more reliable approach)
        $TriggerXml = @"
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(Get-Date -Format "yyyy-MM-dd")T02:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>1</Day>
        </DaysOfMonth>
        <Months>
          <January />
          <February />
          <March />
          <April />
          <May />
          <June />
          <July />
          <August />
          <September />
          <October />
          <November />
          <December />
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
</Task>
"@
        
        # Alternative approach: Create trigger using CIM classes
        $Trigger = New-CimInstance -ClassName MSFT_TaskTrigger -Namespace Root/Microsoft/Windows/TaskScheduler -ClientOnly
        $Trigger.CimInstanceProperties.Item('TriggerType').Value = 3  # TASK_TRIGGER_MONTHLYDATE
        $Trigger.CimInstanceProperties.Item('DaysOfMonth').Value = 1
        $Trigger.CimInstanceProperties.Item('MonthsOfYear').Value = 0xFFF  # All months
        $Trigger.CimInstanceProperties.Item('StartBoundary').Value = (Get-Date -Hour 2 -Minute 0 -Second 0).ToString("yyyy-MM-ddTHH:mm:ss")
        $Trigger.CimInstanceProperties.Item('Enabled').Value = $True
        
        # If CIM approach fails, use simpler weekly approach as fallback
        try {
            # Try to create monthly trigger using Register-ScheduledTask with CIM
            $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $Settings = New-ScheduledTaskSettingsSet `
                -AllowStartIfOnBatteries `
                -DontStopIfGoingOnBatteries `
                -StartWhenAvailable `
                -RestartCount 3 `
                -RestartInterval (New-TimeSpan -Minutes 10) `
                -ExecutionTimeLimit (New-TimeSpan -Hours 2)
            
            # Register task with CIM trigger
            $Task = Register-ScheduledTask `
                -TaskName $TaskName `
                -Action $Action `
                -Trigger $Trigger `
                -Principal $Principal `
                -Settings $Settings `
                -Description "Monthly Exchange Online journaling report generation"
                
        } catch {
            Write-Warning "CIM-based monthly trigger failed, using alternative approach..."
            
            # Fallback: Create a task that runs on the 1st of each month using schtasks.exe
            $TaskCommand = "schtasks.exe"
            $TaskArgs = @(
                "/Create",
                "/TN", "`"$TaskName`"",
                "/TR", "`"PowerShell.exe $ActionArgs`"",
                "/SC", "MONTHLY",
                "/D", "1",
                "/ST", "02:00",
                "/RU", "SYSTEM",
                "/RL", "HIGHEST",
                "/F"
            )
            
            Write-Host "Using schtasks.exe to create monthly task..." -ForegroundColor Cyan
            $Result = & $TaskCommand $TaskArgs
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Monthly task created successfully using schtasks.exe" -ForegroundColor Green
                $Task = Get-ScheduledTask -TaskName $TaskName
            } else {
                throw "schtasks.exe failed with exit code: $LASTEXITCODE"
            }
        }
        
        if ($Task) {
            Write-Host "✓ Scheduled task '$TaskName' created successfully" -ForegroundColor Green
            Write-Host "  - Runs monthly on the 1st at 2:00 AM" -ForegroundColor Cyan
            Write-Host "  - Script: $ScriptPath" -ForegroundColor Cyan
            Write-Host "  - Journal Email: $JournalEmail" -ForegroundColor Cyan
            Write-Host "  - Report Path: $ReportPath" -ForegroundColor Cyan
            return $true
        } else {
            Write-Error "Failed to create scheduled task"
            return $false
        }
        
    } catch {
        Write-Error "Failed to create scheduled task: $($_.Exception.Message)"
        Write-Host "Error Details: $($_.Exception)" -ForegroundColor Red
        
        # Final fallback: Create a simple weekly task
        Write-Host "Attempting fallback: Creating weekly task instead..." -ForegroundColor Yellow
        try {
            $WeeklyTrigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Monday -At "2:00AM"
            $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            $Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -StartWhenAvailable
            
            $FallbackTask = Register-ScheduledTask `
                -TaskName "$TaskName (Weekly Fallback)" `
                -Action $Action `
                -Trigger $WeeklyTrigger `
                -Principal $Principal `
                -Settings $Settings `
                -Description "Weekly Exchange Online journaling report (fallback from monthly)"
            
            if ($FallbackTask) {
                Write-Host "✓ Fallback weekly task created successfully" -ForegroundColor Yellow
                Write-Host "  - Runs every 4 weeks on Monday at 2:00 AM" -ForegroundColor Cyan
                return $true
            }
        } catch {
            Write-Error "Even fallback task creation failed: $($_.Exception.Message)"
        }
        
        return $false
    }
}

# Function to test the scheduled task
function Test-ScheduledTask {
    param([string]$TaskName)
    
    try {
        # Check for exact task name first
        $Task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        
        # If not found, check for fallback task
        if (-not $Task) {
            $Task = Get-ScheduledTask -TaskName "$TaskName (Weekly Fallback)" -ErrorAction SilentlyContinue
        }
        
        if ($Task) {
            Write-Host "✓ Task verification successful" -ForegroundColor Green
            Write-Host "  - Task Name: $($Task.TaskName)" -ForegroundColor Cyan
            Write-Host "  - Task State: $($Task.State)" -ForegroundColor Cyan
            
            # Get additional task info
            $TaskInfo = Get-ScheduledTaskInfo -TaskName $Task.TaskName -ErrorAction SilentlyContinue
            if ($TaskInfo) {
                Write-Host "  - Last Run Time: $($TaskInfo.LastRunTime)" -ForegroundColor Cyan
                Write-Host "  - Next Run Time: $($TaskInfo.NextRunTime)" -ForegroundColor Cyan
                Write-Host "  - Last Task Result: $($TaskInfo.LastTaskResult)" -ForegroundColor Cyan
            }
            return $true
        } else {
            Write-Error "No task found with name '$TaskName' or its fallback variant"
            return $false
        }
    } catch {
        Write-Error "Task verification failed: $($_.Exception.Message)"
        return $false
    }
}

# Function to create manual monthly task using XML
function Create-MonthlyTaskWithXML {
    param(
        [string]$TaskName,
        [string]$ScriptPath,
        [string]$JournalEmail,
        [string]$ReportPath
    )
    
    Write-Host "Creating monthly task using XML definition..." -ForegroundColor Yellow
    
    $ActionArgs = "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -JournalEmailAddress `"$JournalEmail`" -ReportPath `"$ReportPath`""
    
    $TaskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Monthly Exchange Online journaling report generation</Description>
  </RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$(Get-Date -Format "yyyy-MM-dd")T02:00:00</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth>
          <Day>1</Day>
        </DaysOfMonth>
        <Months>
          <January />
          <February />
          <March />
          <April />
          <May />
          <June />
          <July />
          <August />
          <September />
          <October />
          <November />
          <December />
        </Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT2H</ExecutionTimeLimit>
    <Priority>7</Priority>
    <RestartOnFailure>
      <Interval>PT10M</Interval>
      <Count>3</Count>
    </RestartOnFailure>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>PowerShell.exe</Command>
      <Arguments>$ActionArgs</Arguments>
    </Exec>
  </Actions>
</Task>
"@
    
    try {
        # Save XML to temp file
        $TempXMLFile = [System.IO.Path]::GetTempFileName() + ".xml"
        $TaskXML | Out-File -FilePath $TempXMLFile -Encoding Unicode
        
        # Import task using schtasks
        $Result = schtasks.exe /Create /TN "$TaskName" /XML "$TempXMLFile" /F
        
        # Clean up temp file
        Remove-Item $TempXMLFile -Force -ErrorAction SilentlyContinue
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Monthly task created successfully using XML definition" -ForegroundColor Green
            return $true
        } else {
            Write-Warning "XML-based task creation failed with exit code: $LASTEXITCODE"
            return $false
        }
        
    } catch {
        Write-Error "XML-based task creation failed: $($_.Exception.Message)"
        return $false
    }
}

# Main execution
Write-Host "=== Creating Scheduled Task for Exchange Online Reports ===" -ForegroundColor Cyan
Write-Host "Task Name: $TaskName" -ForegroundColor White
Write-Host "Script Path: $ScriptPath" -ForegroundColor White
Write-Host "Journal Email: $JournalEmailAddress" -ForegroundColor White
Write-Host "Report Path: $ReportPath" -ForegroundColor White
Write-Host ""

# Validate prerequisites
Write-Host "Validating prerequisites..." -ForegroundColor Yellow

# Check if script file exists
if (-not (Test-Path $ScriptPath)) {
    Write-Error "Script file not found at: $ScriptPath"
    Write-Host "Please ensure the Configure-ExchangeJournaling.ps1 script is saved to this location." -ForegroundColor Red
    exit 1
}

# Check if running as administrator
$CurrentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = New-Object Security.Principal.WindowsPrincipal($CurrentUser)
$IsAdmin = $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $IsAdmin) {
    Write-Warning "This script should be run as Administrator for best results."
    Write-Host "Some features may not work properly without elevated privileges." -ForegroundColor Yellow
}

# Create report directory if it doesn't exist
if (-not (Test-Path $ReportPath)) {
    try {
        New-Item -ItemType Directory -Path $ReportPath -Force | Out-Null
        Write-Host "✓ Created report directory: $ReportPath" -ForegroundColor Green
    } catch {
        Write-Error "Failed to create report directory: $($_.Exception.Message)"
        exit 1
    }
}

# Try multiple approaches to create the monthly task
Write-Host "`nCreating scheduled task..." -ForegroundColor Yellow

$TaskCreated = $false

# Approach 1: Try XML-based creation first
if (Create-MonthlyTaskWithXML -TaskName $TaskName -ScriptPath $ScriptPath -JournalEmail $JournalEmailAddress -ReportPath $ReportPath) {
    $TaskCreated = $true
}

# Approach 2: Try PowerShell cmdlets if XML failed
if (-not $TaskCreated) {
    Write-Host "Trying PowerShell cmdlet approach..." -ForegroundColor Yellow
    if (Create-ScheduledTask -TaskName $TaskName -ScriptPath $ScriptPath -JournalEmail $JournalEmailAddress -ReportPath $ReportPath) {
        $TaskCreated = $true
    }
}

if ($TaskCreated) {
    # Test the scheduled task
    Write-Host "`nTesting scheduled task..." -ForegroundColor Yellow
    if (Test-ScheduledTask -TaskName $TaskName) {
        Write-Host "`n✓ Setup completed successfully!" -ForegroundColor Green
        Write-Host "The system will now automatically generate monthly reports." -ForegroundColor Green
        
        # Provide additional information
        Write-Host "`n--- Additional Information ---" -ForegroundColor Cyan
        Write-Host "• To run the task manually: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor White
        Write-Host "• To view task history: Get-ScheduledTaskInfo -TaskName '$TaskName'" -ForegroundColor White
        Write-Host "• To modify the task: Use Task Scheduler (taskschd.msc)" -ForegroundColor White
        Write-Host "• Reports will be saved to: $ReportPath" -ForegroundColor White
        
        # Ask if user wants to run a test
        $RunTest = Read-Host "`nWould you like to run a test execution now? (y/n)"
        if ($RunTest -eq 'y' -or $RunTest -eq 'Y') {
            Write-Host "`nStarting test execution..." -ForegroundColor Yellow
            try {
                Start-ScheduledTask -TaskName $TaskName
                Write-Host "✓ Test execution started. Check the report directory and Windows Event Log for results." -ForegroundColor Green
            } catch {
                Write-Warning "Failed to start test execution: $($_.Exception.Message)"
            }
        }
        
    } else {
        Write-Error "Task creation succeeded but verification failed."
        exit 1
    }
    
} else {
    Write-Error "Failed to create scheduled task using all available methods."
    Write-Host "`nManual Alternative:" -ForegroundColor Yellow
    Write-Host "You can manually create the task using Task Scheduler:" -ForegroundColor White
    Write-Host "1. Open Task Scheduler (taskschd.msc)" -ForegroundColor White
    Write-Host "2. Create Basic Task" -ForegroundColor White
    Write-Host "3. Set trigger to Monthly, 1st day, 2:00 AM" -ForegroundColor White
    Write-Host "4. Set action to start PowerShell.exe with arguments:" -ForegroundColor White
    Write-Host "   -ExecutionPolicy Bypass -File `"$ScriptPath`" -JournalEmailAddress `"$JournalEmailAddress`" -ReportPath `"$ReportPath`"" -ForegroundColor Cyan
    exit 1
}

Write-Host "`n=== Script Execution Completed ===" -ForegroundColor Cyan
