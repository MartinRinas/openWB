  
<#
  .SYNOPSIS
  Creates ad-hoc and scheduled backup of openWB wallbox.
  .DESCRIPTION
    This script will create a backup of the openWB wallbox and store the backup in a folder while keeping older versions. The script can also register a scheduled task to automate the backup creation. 
    Starting the script without any parameters will 1) ask for the IP of the openWB, 2) check if a scheduled task already exists. A new task will be registered if needed and an immediate backup will be created. Script will not perform any action if backup job already exists. 
    Start script with -RunOnce parameter to perform an immediate backup without creating the scheduled task.
  .PARAMETER LocalBackupFolder
  Specifies local path to store backups. Defaults to script location
  .PARAMETER OpenWBIP
  IP of openWB, required parameter.
  .PARAMETER VerboseLogFileName
  Logfile for Verbose output, defaults to VerboseOutput.Log if not specified
  .PARAMETER AppendToLogfile
  Appends to existing verbose log file, defaults to override.
  .PARAMETER RunOnce
  Execute backup directly, skip creation of scheduled task.
    .EXAMPLE
  PS> .\Get-OpenWBBackup.ps1
  Checks for scheduled task, creates one if it doesn't exist already. openWB backup will be executed every 4 weeks, Monday at 09:00 local time. Reruns if time is missed. After creating the task an immediate backup is created. 
  Skips all actions if a scheduled task already exists.
  .EXAMPLE
  PS> .\Get-OpenWBBackup.ps1 -OpenWBIP 192.168.178.200 -RunOnce
  Performs an immediate backup of OpenWB 192.168.178.200 and skips creation of scheduled task.
  
  .NOTES
  MIT License
  Copyright (c) 2021 Martin Rinas
  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:
  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.
  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.
#>
[CmdletBinding()] #few examples how a script accepts input parameters
param
(
    [Parameter(Mandatory=$false)]
    [ValidateScript({Test-Path $_ })]
    [string]$LocalBackupFolder,
    [Parameter(Mandatory=$true,HelpMessage="Please enter IP of openWB")]
    [string]$OpenWBIP,
    [Parameter(Mandatory=$false)]
    [string] $VerboseLogFileName = 'VerboseOutput.Log',
    [switch]$AppendToLogfile,
    [switch]$RunOnce
)



#region function definition
function Write-VerboseLog{
    <#
        .SYNOPSIS
        Writes messages to Verbose output and Verbose log file
        .DESCRIPTION
        This fuction will direct verbose output to the console if the -Verbose 
        paramerter is specified.  It will also write output to the Verboselog.
        .EXAMPLE
        Write-VerboseLog -Message 'This is a single verbose message'
        .EXAMPLE
        # Write-VerboseLog -Message ('This is a more complex version where we want to include the value of a variable: {0}' -f ($MyVariable | Out-String) )
    #>
    param(
      [String]$Message  
    )
  
    $VerboseMessage = ('{0} Line:{1} {2}' -f (Get-Date), $MyInvocation.ScriptLineNumber, $Message)
    Write-Verbose -Message $VerboseMessage
    Add-Content -Path $VerboseLogFileName -Value $VerboseMessage
  }
function Register-OpenWBBackupTask
{
    param
    (
        [Parameter(Mandatory=$true,HelpMessage='Full path to the backup script')]
        [ValidateScript({Test-Path $_ })]
        [String]$ScriptPath,
        [Parameter(Mandatory=$true,HelpMessage='Name of backuptask')]
        [string]$TaskName,
        [Parameter(Mandatory=$true,HelpMessage='IP of OpenWB')]
        [string]$IPAddr

    )
    try 
    {
        $null = Get-ScheduledTask -TaskName $TaskName -ErrorAction stop
        Write-VerboseLog -Message ('Scheduled Task: {0} does already exist. Please delete first if you want to re-create' -f ($TaskName | Out-String) )
        Write-Error "Task $TaskName already exists. Will not re-create, please delete manually in Task Scheduler if you want to re-create. Exiting."
        break
    }
    catch 
    {
        # Scheduled Task does not exist. Creating.
        Write-VerboseLog -Message ('Scheduled task: {0} does not exist, creating.'-f ($TaskName | Out-String) )
        $WorkingDir = Split-Path $ScriptPath -Parent
        Write-VerboseLog -Message ('Setting Working directory to: {0}' -f($WorkingDir | out-string))
        $Action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-ExecutionPolicy Bypass -File  `"$ScriptPath`" -RunOnce -openWBIP $IPAddr" -WorkingDirectory $WorkingDir
        $Trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -At '09:00' -DaysOfWeek Monday
        $Principal = New-ScheduledTaskPrincipal -LogonType Interactive -Id Author -UserId $env:USERNAME -RunLevel Limited -ProcessTokenSidType Default
        $Settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Hours 1) -StartWhenAvailable -DontStopIfGoingOnBatteries -AllowStartIfOnBatteries
        $ScheduledTask = New-ScheduledTask -Action $Action -Principal $Principal -Settings $Settings -Trigger $Trigger -Description 'OpenWB BackupJob'
        Register-ScheduledTask $TaskName -InputObject $ScheduledTask -ErrorAction Stop
    }

}

#endregion

#region variable definition
if ([string]::IsNullOrEmpty($LocalBackupFolder))
{
    $LocalBackupFolder = (Split-path $MyInvocation.MyCommand.Path -parent)
    Write-VerboseLog -Message ('LocalBackupFolder was empty, set to {0}' -f ($LocalBackupFolder))
}

if([string]::IsNullOrEmpty($ScriptLocation))
{
    $ScriptLocation = (Split-path $MyInvocation.MyCommand.Path -parent)
    Write-VerboseLog -Message ('ScriptLocation was empty, set to {0}' -f ($ScriptLocation))
}

$LocalBackupFileName = ("OpenWB-backup-" + (get-date -format u) + ".tar.gz").Replace(":","-")
$LocalBackupPath = $LocalBackupFolder + '\' + $LocalBackupFileName
$OpenWBBackupPath = '/openWB/web/settings/backup.php'
$URIToCall = "http://" + $OpenWBIP + $OpenWBBackupPath
$BackupTaskName = 'openWB Backup'
$ScriptName = $MyInvocation.MyCommand

#endregion

#region main script
if (!$AppendToLogfile)
{
    $null = New-Item -Path $VerboseLogFileName -ItemType File -Force
    Write-VerboseLog -Message 'Script started'
    Write-VerboseLog -Message ('Parameters: {0}' -f ($MyInvocation | Out-String) )
}
else 
{
    Write-VerboseLog -Message 'Script started, appending to existing log'
}

if ($RunOnce)
{
    Write-VerboseLog 'Skip scheduled task and executy directly'
}
else
{
    # check for and create scheduled task
    Write-VerboseLog -message ('Checking for existence of scheduled task: {0} and creating if needed. ' -f ($BackupTaskName) )
    
    try 
    {
        $null = Get-ScheduledTask -TaskName $BackupTaskName -ErrorAction stop
        Write-VerboseLog -Message ('Scheduled Task: {0} does already exist. ' -f ($BackupTaskName) )
        $BackupTaskExists = $True
    }
    catch 
    {
        Write-VerboseLog -Message ('Scheduled Task: {0} does not exist. ' -f ($BackupTaskName) )
        $BackupTaskExists = $False
    }

    if ($BackupTaskExists)
    {
        Write-VerboseLog -Message ('nothing to do.' -f ($BackupTaskName) )
        Write-Host "Backupjob $BackupTaskName already exists, nothing to do. Exiting..."
        exit
    }
    else
    {
        Write-VerboseLog -Message ('BackupTask: {0} does not already exist, creating.' -f ($BackupTaskName) )    
        Write-Host "Creating scheduled backup task for openWB"
        $FullScriptPath = $ScriptLocation + '\' + $ScriptName
        Write-VerboseLog -Message ('Full script location: {0}' -f ($FullScriptPath| Out-String))
        Register-OpenWBBackupTask -TaskName $BackupTaskName -ScriptPath $FullScriptPath -IPAddr $OpenWBIP
        Write-Host "Creating one backup immediately."
    }
}

Write-VerboseLog -Message ('Using this as backup path to store file: {0}' -f ($LocalBackupPath | Out-String) )

Write-Host "Starting backup of openWB"
Write-VerboseLog -Message ('Triggering openWB backup creation by calling: {0}' -f ($URIToCall | Out-String) )
$Result = Invoke-WebRequest -uri $URIToCall -UseBasicParsing  #create backup

if ($Result.StatusCode -eq '200') # New backup created?
{
    try 
    {
        #$OpenWBBackupDownloadPath = '/openWB/web/backup/backup.tar.gz' # We could also use $Result.Links.href to dynamically fetch location if we wanted to. Does require extra handling if more than one link is provided
        if ($Result.Links.Count -gt 1)
        {
            Write-VerboseLog -Message ('More than one link found in response: {0}' -f ($result.links | Out-String) )
            throw "More than one link found in response, cannot proceed."
        }
        $OpenWBBackupDownloadPath = $Result.Links.href
        $BackupURI = "http://" + $OpenWBIP + $OpenWBBackupDownloadPath
        Write-VerboseLog -Message ( 'Backup generated, downloading from: {0} to : {1}' -f ($BackupURI | Out-String), ($LocalBackupPath | Out-String) )
        Invoke-WebRequest -Uri $BackupURI -OutFile $LocalBackupPath # Downlaod backup and store locally
        Write-Host "Created backup at $LocalBackupPath" # we're done here
        Write-VerboseLog -Message 'Download completed.'
    }  
    catch 
    {
        Write-VerboseLog 'Backup created bout could not be downloaded'
        Write-Host "Backup created but couldn't be downloaded." 
    }
}
else 
{
    Write-VerboseLog -Message ( 'Unexpected return code when asking for backup: {0} {1}' -f ($result.StatusCode | Out-String), ($Result.StatusDescription | Out-String ) )
    Write-Host 'Unexpected return code when asking for backup:' ($Result.StatusCode) ($Result.StatusDescription)   
}
#endregion