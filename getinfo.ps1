function Get-AccessRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,                            # Path to .accdb or .mdb
        [Parameter(Mandatory)]
        [string]$Query,                           # e.g., "SELECT * FROM Customers"
        [string]$Password,                        # Optional database password
        [ValidateSet('12.0','16.0')]              # 12.0 (2010) or 16.0 (2016+)
        [string]$AceVersion = '12.0'
    )

    # Build connection string
    $connString = "Provider=Microsoft.ACE.OLEDB.$AceVersion;Data Source=$Path;"
    if ($Password) { $connString += "Jet OLEDB:Database Password=$Password;" }

    # Load .NET types
    Add-Type -AssemblyName System.Data

    $connection = New-Object System.Data.OleDb.OleDbConnection($connString)
    try {
        $connection.Open()

        $command = $connection.CreateCommand()
        $command.CommandText = $Query

        $adapter = New-Object System.Data.OleDb.OleDbDataAdapter($command)
        $dataSet = New-Object System.Data.DataSet
        [void]$adapter.Fill($dataSet)

        # Convert rows to PowerShell objects
        $table = $dataSet.Tables[0]
        foreach ($row in $table.Rows) {
            $obj = [pscustomobject]@{}
            foreach ($col in $table.Columns) {
                $obj | Add-Member -NotePropertyName $col.ColumnName -NotePropertyValue $row[$col.ColumnName]
            }
            $obj
        }
    }
    catch {
        Write-Error "Failed to query Access DB: $($_.Exception.Message)"
    }
    finally {
        if ($connection.State -ne 'Closed') { $connection.Close() }
        $connection.Dispose()
    }
}

# EXAMPLE: read entire table
#Get-AccessRows -Path "c:\PSCommand.accdb" -Query "SELECT * FROM Test" -AceVersion 16.0

# EXAMPLE: filtered query
#Get-AccessRows -Path "c:\PSCommand.accdb" -Query "SELECT * FROM [TEST] WHERE [Last Name]"

Get-AccessRows -Path "c:test.accdb" -Query "SELECT * FROM [New Year]"

#Last time the file was modified 
#Get-AccessRows -Path "c:test.accdb" -Query "SELECT * FROM [New Year] WHERE [Last Modified] >= #2024-01-01#"

#OrderID, CustomerID, OrderDate FROM Orders WHERE OrderDate >= #2024-01-01#"    WHERE [Last Modified] >= #2024-01-01#