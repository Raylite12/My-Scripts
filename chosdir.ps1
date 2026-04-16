show-command

$fold = Get-ChildItem -Directory | Out-GridView -Title 'Pick a folder' -PassThru; if ($fold) {Set-Location $fold.FullName}

Get-Command -Verb Get