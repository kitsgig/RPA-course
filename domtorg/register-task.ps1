param(
    [string] $Python = "$PSScriptRoot\.venv\Scripts\python.exe",
    [string] $Config = "$PSScriptRoot\config.toml",
    [string] $TaskName = "DomTorg catalog",
    [switch] $XmlOnly
)
$ErrorActionPreference = "Stop"

$pythonPath = (Resolve-Path -LiteralPath $Python).Path
$configPath = (Resolve-Path -LiteralPath $Config).Path
$robotPath = (Resolve-Path -LiteralPath "$PSScriptRoot\robot.py").Path
$interval = & $pythonPath $robotPath --config $configPath --interval
if ($LASTEXITCODE -ne 0 -or $interval -notmatch '^[1-9][0-9]*$' -or [int]$interval -gt 540) {
    throw "Check robot.interval_minutes in config.toml (1-540)."
}

$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$user = [System.Security.SecurityElement]::Escape($identity)
$command = [System.Security.SecurityElement]::Escape($pythonPath)
$arguments = [System.Security.SecurityElement]::Escape(('"{0}" --config "{1}" --scheduled' -f $robotPath, $configPath))
$workdir = [System.Security.SecurityElement]::Escape($PSScriptRoot)
$start = [DateTimeOffset]::UtcNow.ToOffset([TimeSpan]::FromHours(8)).ToString("yyyy-MM-dd") + "T09:00:00+08:00"
$repeat = "PT{0}M" -f $interval
$duration = "PT{0}M" -f (540 + [int]$interval)
$xml = @"
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$start</StartBoundary>
      <Repetition>
        <Interval>$repeat</Interval>
        <Duration>$duration</Duration>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <ScheduleByDay><DaysInterval>1</DaysInterval></ScheduleByDay>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Robot">
      <UserId>$user</UserId>
      <LogonType>Password</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
  </Settings>
  <Actions Context="Robot">
    <Exec>
      <Command>$command</Command>
      <Arguments>$arguments</Arguments>
      <WorkingDirectory>$workdir</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

if ($XmlOnly) {
    $xml
    return
}

$credential = Get-Credential -UserName $identity -Message "Enter the Windows password for the robot account."
if (-not $credential -or $credential.UserName -ne $identity) {
    throw "Register the task as the current Windows account."
}
$task = @{
    TaskName = $TaskName
    Xml = $xml
    User = $identity
    Password = $credential.GetNetworkCredential().Password
}
Register-ScheduledTask @task
