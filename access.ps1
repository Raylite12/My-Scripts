$Access = New-Object -ComObject Access.Application
$Access.Visible = $true

$dbpath = "c:\PSCommand.accdb"
$Access.OpenCurrentDatabase($dbpath)

$tableName = "Command"

$Access.DoCmd.OpenTable($tableName)

$connection = New-Object -ComObject ADODB.Connection
$connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbpath"
$connection.Open($connectionString)

$recordset = New-Object -ComObject ADODB.RecordSet 

$recordset.LockType = 3
$recordset.CursorType = 3

start-sleep -Seconds 10
$recordset.Open("SELECT * FROM [$tableName]", $connection)

$recordset.AddNew()
$recordset.Fields("Description").Value = "Test"
$recordset.Update()

$recordset.Close()
$connection.Close()
