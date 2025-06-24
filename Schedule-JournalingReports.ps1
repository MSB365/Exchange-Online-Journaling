#region Description
<#     
.NOTES
==============================================================================
Created on:         2025/07/01
Created by:         Drago Petrovic
Organization:       MSB365.blog
Filename:           Schedule-JournalingReports.ps1
Current version:    V1.0     

Find us on:
* Website:         https://www.msb365.blog
* Technet:         https://social.technet.microsoft.com/Profile/MSB365
* LinkedIn:        https://www.linkedin.com/in/drago-petrovic/
* MVP Profile:     https://mvp.microsoft.com/de-de/PublicProfile/5003446
==============================================================================

.SYNOPSIS
    Manages Exchange Online mailbox permissions based on CSV file input with GUI file picker and HTML reporting

.DESCRIPTION
    This script sets "Send As" and "Full Access" permissions for mailboxes defined in a CSV file.
    It removes any existing permissions that are not defined in the CSV file.
    Features a GUI file picker for easy CSV file selection and generates detailed HTML reports.

.EXAMPLE
	# Set up automated monthly reports
	.\Schedule-JournalingReports.ps1 -JournalEmailAddress "journal@yourdomain.com"

    

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
===========================================================================
.CHANGE LOG
V1.00, 2025/07/01 - DrPe - Initial version



--- keep it simple, but significant ---


--- by MSB365 Blog ---

#>
#endregion
##############################################################################################################
[cmdletbinding()]
param(
[switch]$accepteula,
[switch]$v)

###############################################################################
#Script Name variable
$Scriptname = "Schedule JournalingReports EXO"
$RKEY = "MSB365_Schedule-JournalingReports"
###############################################################################

[void][System.Reflection.Assembly]::Load('System.Drawing, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a')
[void][System.Reflection.Assembly]::Load('System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089')

function ShowEULAPopup($mode)
{
    $EULA = New-Object -TypeName System.Windows.Forms.Form
    $richTextBox1 = New-Object System.Windows.Forms.RichTextBox
    $btnAcknowledge = New-Object System.Windows.Forms.Button
    $btnCancel = New-Object System.Windows.Forms.Button

    $EULA.SuspendLayout()
    $EULA.Name = "MIT"
    $EULA.Text = "$Scriptname - License Agreement"

    $richTextBox1.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $richTextBox1.Location = New-Object System.Drawing.Point(12,12)
    $richTextBox1.Name = "richTextBox1"
    $richTextBox1.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $richTextBox1.Size = New-Object System.Drawing.Size(776, 397)
    $richTextBox1.TabIndex = 0
    $richTextBox1.ReadOnly=$True
    $richTextBox1.Add_LinkClicked({Start-Process -FilePath $_.LinkText})
    $richTextBox1.Rtf = @"
{\rtf1\ansi\ansicpg1252\deff0\nouicompat{\fonttbl{\f0\fswiss\fprq2\fcharset0 Segoe UI;}{\f1\fnil\fcharset0 Calibri;}{\f2\fnil\fcharset0 Microsoft Sans Serif;}}
{\colortbl ;\red0\green0\blue255;}
{\*\generator Riched20 10.0.19041}{\*\mmathPr\mdispDef1\mwrapIndent1440 }\viewkind4\uc1
\pard\widctlpar\f0\fs19\lang1033 MSB365 SOFTWARE MIT LICENSE\par
Copyright (c) 2025 Drago Petrovic\par
$Scriptname \par
\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}These license terms are an agreement between you and MSB365 (or one of its affiliates). IF YOU COMPLY WITH THESE LICENSE TERMS, YOU HAVE THE RIGHTS BELOW. BY USING THE SOFTWARE, YOU ACCEPT THESE TERMS.\par
\par
MIT License\par
{\pict{\*\picprop}\wmetafile8\picw26\pich26\picwgoal32000\pichgoal15
0100090000035000000000002700000000000400000003010800050000000b0200000000050000
000c0202000200030000001e000400000007010400040000000701040027000000410b2000cc00
010001000000000001000100000000002800000001000000010000000100010000000000000000
000000000000000000000000000000000000000000ffffff00000000ff040000002701ffff0300
00000000
}\par
\pard
{\pntext\f0 1.\tab}{\*\pn\pnlvlbody\pnf0\pnindent0\pnstart1\pndec{\pntxta.}}
\fi-360\li360 Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \par
\pard\widctlpar\par
\pard\widctlpar\li360 The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
\pard\widctlpar\fi-360\li360 2.\tab THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 3.\tab IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. \par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360 4.\tab DISCLAIMER OF WARRANTY. THE SOFTWARE IS PROVIDED \ldblquote AS IS,\rdblquote  WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL MSB365 OR ITS LICENSORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THE SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 5.\tab LIMITATION ON AND EXCLUSION OF DAMAGES. IF YOU HAVE ANY BASIS FOR RECOVERING DAMAGES DESPITE THE PRECEDING DISCLAIMER OF WARRANTY, YOU CAN RECOVER FROM MICROSOFT AND ITS SUPPLIERS ONLY DIRECT DAMAGES UP TO U.S. $1.00. YOU CANNOT RECOVER ANY OTHER DAMAGES, INCLUDING CONSEQUENTIAL, LOST PROFITS, SPECIAL, INDIRECT, OR INCIDENTAL DAMAGES. This limitation applies to (i) anything related to the Software, services, content (including code) on third party Internet sites, or third party applications; and (ii) claims for breach of contract, warranty, guarantee, or condition; strict liability, negligence, or other tort; or any other claim; in each case to the extent permitted by applicable law. It also applies even if MSB365 knew or should have known about the possibility of the damages. The above limitation or exclusion may not apply to you because your state, province, or country may not allow the exclusion or limitation of incidental, consequential, or other damages.\par
\pard\widctlpar\par
\pard\widctlpar\fi-360\li360\qj 6.\tab ENTIRE AGREEMENT. This agreement, and any other terms MSB365 may provide for supplements, updates, or third-party applications, is the entire agreement for the software.\par
\pard\widctlpar\qj\par
\pard\widctlpar\fi-360\li360\qj 7.\tab A complete script documentation can be found on the website https://www.msb365.blog.\par
\pard\widctlpar\par
\pard\sa200\sl276\slmult1\f1\fs22\lang9\par
\pard\f2\fs17\lang2057\par
}
"@
    $richTextBox1.BackColor = [System.Drawing.Color]::White
    $btnAcknowledge.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnAcknowledge.Location = New-Object System.Drawing.Point(544, 415)
    $btnAcknowledge.Name = "btnAcknowledge";
    $btnAcknowledge.Size = New-Object System.Drawing.Size(119, 23)
    $btnAcknowledge.TabIndex = 1
    $btnAcknowledge.Text = "Accept"
    $btnAcknowledge.UseVisualStyleBackColor = $True
    $btnAcknowledge.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::Yes})

    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.Location = New-Object System.Drawing.Point(669, 415)
    $btnCancel.Name = "btnCancel"
    $btnCancel.Size = New-Object System.Drawing.Size(119, 23)
    $btnCancel.TabIndex = 2
    if($mode -ne 0)
    {
   $btnCancel.Text = "Close"
    }
    else
    {
   $btnCancel.Text = "Decline"
    }
    $btnCancel.UseVisualStyleBackColor = $True
    $btnCancel.Add_Click({$EULA.DialogResult=[System.Windows.Forms.DialogResult]::No})

    $EULA.AutoScaleDimensions = New-Object System.Drawing.SizeF(6.0, 13.0)
    $EULA.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
    $EULA.ClientSize = New-Object System.Drawing.Size(800, 450)
    $EULA.Controls.Add($btnCancel)
    $EULA.Controls.Add($richTextBox1)
    if($mode -ne 0)
    {
   $EULA.AcceptButton=$btnCancel
    }
    else
    {
        $EULA.Controls.Add($btnAcknowledge)
   $EULA.AcceptButton=$btnAcknowledge
        $EULA.CancelButton=$btnCancel
    }
    $EULA.ResumeLayout($false)
    $EULA.Size = New-Object System.Drawing.Size(800, 650)

    Return ($EULA.ShowDialog())
}

function ShowEULAIfNeeded($toolName, $mode)
{
$eulaRegPath = "HKCU:Software\Microsoft\$RKEY"
$eulaAccepted = "No"
$eulaValue = $toolName + " EULA Accepted"
if(Test-Path $eulaRegPath)
{
$eulaRegKey = Get-Item $eulaRegPath
$eulaAccepted = $eulaRegKey.GetValue($eulaValue, "No")
}
else
{
$eulaRegKey = New-Item $eulaRegPath
}
if($mode -eq 2) # silent accept
{
$eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
else
{
if($eulaAccepted -eq "No")
{
$eulaAccepted = ShowEULAPopup($mode)
if($eulaAccepted -eq [System.Windows.Forms.DialogResult]::Yes)
{
        $eulaAccepted = "Yes"
        $ignore = New-ItemProperty -Path $eulaRegPath -Name $eulaValue -Value $eulaAccepted -PropertyType String -Force
}
}
}
return $eulaAccepted
}

if ($accepteula)
    {
         ShowEULAIfNeeded "DS Authentication Scripts:" 2
         "EULA Accepted"
    }
else
    {
        $eulaAccepted = ShowEULAIfNeeded "DS Authentication Scripts:" 0
        if($eulaAccepted -ne "Yes")
            {
                "EULA Declined"
                exit
            }
         "EULA Accepted"
    }
###############################################################################
write-host "  _           __  __ ___ ___   ____  __ ___  " -ForegroundColor Yellow
write-host " | |__ _  _  |  \/  / __| _ ) |__ / / /| __| " -ForegroundColor Yellow
write-host " | '_ \ || | | |\/| \__ \ _ \  |_ \/ _ \__ \ " -ForegroundColor Yellow
write-host " |_.__/\_, | |_|  |_|___/___/ |___/\___/___/ " -ForegroundColor Yellow
write-host "       |__/                                  " -ForegroundColor Yellow
Start-Sleep -s 2
write-host ""                                                                                   
write-host ""
write-host ""
write-host ""
###############################################################################


#----------------------------------------------------------------------------------------
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
