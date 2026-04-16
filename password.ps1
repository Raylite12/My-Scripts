[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#A function to make random password
function New-RandomPassword {
param([int]$length = 20)
    $characterSet = @{
	    Lowercase   = (97..122) | Get-Random -Count $length | ForEach-Object { [Char]$_ }
		Uppercase   = (65..90)  | Get-Random -Count $length | ForEach-Object { [Char]$_ }
		Numeric     = (48..57)  | Get-Random -Count $length | ForEach-Object { [Char]$_ }
		SpecialChar = (33..47) + (58..64) + (91..96) + (123..126) | Get-Random -Count $length | ForEach-Object { [Char]$_ }
	}
	$StringSet = $CharacterSet.Uppercase + $CharacterSet.Lowercase + $CharacterSet.Numeric + $CharacterSet.SpecialChar
	-join (Get-Random -Count $length -InputObject $StringSet)
}
New-RandomPassword

#Ask for input
param([Parameter(Mandatory)][string]$fullName)

$action = New-ScheduledTaskAction -Execute "Taskmgr.exe"
$trigger = New-ScheduledTaskTrigger -AtLogon
$principal = "Contoso\Administrator"
$settings = New-ScheduledTaskSettingsSet
$task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings
Register-ScheduledTask T1 -InputObject $task

<# In this example, the set of commands uses several cmdlets and variables to define and then register a scheduled task.

The first command uses the New-ScheduledTaskAction cmdlet to assign the executable file tskmgr.exe to the variable $action.

The second command uses the New-ScheduledTaskTrigger cmdlet to assign the value AtLogon to the variable $trigger.

The third command assigns the principal of the scheduled task Contoso\Administrator to the variable $principal.

The fourth command uses the New-ScheduledTaskSettingsSet cmdlet to assign a task settings object to the variable $settings.

The fifth command creates a new task and assigns the task definition to the variable $task.

The sixth command (hypothetically) runs at a later time. It registers the new scheduled task and defines it by 
using the $task variable. #>
