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