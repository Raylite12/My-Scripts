[net.servicePointManager]::SecurityProtocol = 
[Net.ServicePointManager]::SecurityProtocol -Bor 
[Net.SecurityProtocolType]::Tls12

Connect-AzAccount -Tenant
Connect-MgGraph
Get-MgContext | fl *
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

Get-MgGroup | Format-List DisplayName | Out-File -FilePath c: -Append
Get-MgUser | Format-List DisplayName | Out-File -FilePath c: -Append

function commandlr {
    #Define the path to your input file 
    $inputFile = "c:command1.txt"

    #Define the path for the output Excel file
    $outputTxt = "c:Larry1"

    #Get the content of the text file
    $textContext = Get-Content $inputFile 

    #Create an Excel Com Object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true       #Make Excel visible (remove or set false for background operation)

    #Create a new workbook
    $workbook = $excel.workbooks.Add()

    #Get the first worksheet 
    #Access the first default sheet and rename it
    $sheetGroup = $workbook.Sheets.Item(1)   #Access the first sheet by index 
    $sheetGroup.Name = "Command"

    #Add a title to the worksheet
    $sheetGroup.Cells.Item(1, 1).value = "Command Type"
    $sheetGroup.Cells.Item(1, 1).Font.Bold = $true
    $sheetGroup.Cells.Item(1, 1).Font.Size = 12
    $sheetGroup.Cells.Item(1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)

    $sheetGroup.Cells.Item(1, 2).value = "Name"
    $sheetGroup.Cells.Item(1, 2).Font.Bold = $true
    $sheetGroup.Cells.Item(1, 2).Font.Size = 12
    $sheetGroup.Cells.Item(1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BurlyWood)

    #Add a second sheet and rename it 
    $sheetUser = $workbook.Sheets.add()
    $sheetUser.Name = "User"

    $sheetUser.Cells.Item(1, 1).value = "Command Type"
    $sheetUser.Cells.Item(1, 1).Font.Bold = $true
    $sheetUser.Cells.Item(1, 1).Font.Size = 12
    $sheetUser.Cells.Item(1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)

    $sheetUser.Cells.Item(1, 2).value = "Name"
    $sheetUser.Cells.Item(1, 2).Font.Bold = $true
    $sheetUser.Cells.Item(1, 2).Font.Size = 12
    $sheetUser.Cells.Item(1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BurlyWood)

    #Function to set the order of sheets
    function Set-SheetOrder {
        param (
            $workbook,
            [string[]]$order
        )
       # $sheetCount = $workbook.Sheets.$sheet.Count
       $index = 1
        foreach ($name in $order) {
            $sheet = $workbook.Sheets.Item($name)
            $sheet.Move($workbook.Sheets.Item($index))
            $index++
        }
        
    }
    #Set the order of the sheets
    Set-SheetOrder -Workbook $workbook -Order @("Group", "Users")

   #Populate the Excel sheet with data from the text file    
for ($i = 0; $i -lt $textContext.Length; $i++) {
    $sheetGroup.cells.Item($i + 2, 1).Value = $textContext[$i]

    $columns = $textContext[$i].Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    for ($j = 0; $j -lt $columns.Length; $j++) {
        $sheetGroup.cells.item($i + 2, $j + 1).Value = $columns[$j]
    }
}

#Autofit and Font Name the columns
$sheetGroup.UsedRange.Font.Name = "Times New Roman"
$sheetGroup.UsedRange.Font.Size = 12
$sheetGroup.UsedRange.Cells.EntireColumn.AutoFit()

#Save the workbook
$workbook.SaveAs($outputTxt)

#Clean up
<#for ($i = 10; $i -ge 1; $i--) {
    Write-Host "Time left: $i seconds"
    Start-Sleep -Seconds 1
}
Write-Host "Done!"#>
function Start-Stop {
    param(
        [int]$Seconds
    )

    Write-Host "Starting timer for $Seconds seconds..."
    Start-Sleep -Seconds $Seconds
    Write-Host "Timer finished."
}

Start-Stop -Seconds 10
$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

}
commandlr
