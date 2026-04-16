Get-Command | Out-File -FilePath "c:\command1.txt"

function commandlr {
    #Define the path to your input file 
    $inputFile = "c:\command1.txt"

    #Define the path for the output Excel file
    $outputTxt = "c:test"

    #Get the content of the text file
    $textContext = Get-Content $inputFile 

    #Create an Excel Com Object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true       #Make Excel visible (remove or set false for background operation)

    #Create a new workbook
    $workbook = $excel.workbooks.Add()

    #Get the first worksheet 
    #Access the first default sheet and rename it
    $sheet = $workbook.Sheets.Item(1)   #Access the first sheet by index 
    $sheet.Name = "Command"

    #Add a title to the worksheet
    $sheet.Cells.Item(1, 1).value = "Command Type"
    $sheet.Cells.Item(1, 1).Font.Bold = $true
    $sheet.Cells.Item(1, 1).Font.Size = 12
    $sheet.Cells.Item(1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)

    $sheet.Cells.Item(1, 2).value = "Name"
    $sheet.Cells.Item(1, 2).Font.Bold = $true
    $sheet.Cells.Item(1, 2).Font.Size = 12
    $sheet.Cells.Item(1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BurlyWood)

    $sheet.Cells.Item(1, 3).value = "Version"
    $sheet.Cells.Item(1, 3).Font.Bold = $true
    $sheet.Cells.Item(1, 3).Font.Size = 12
    $sheet.Cells.Item(1, 3).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::RoyalBlue)

    $sheet.Cells.Item(1, 4).value = "Source"
    $sheet.Cells.Item(1, 4).Font.Bold = $true
    $sheet.Cells.Item(1, 4).Font.Size = 12
    $sheet.Cells.Item(1, 4).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightYellow)

    $sheet.Cells.Item(1, 5).value = "Description"
    $sheet.Cells.Item(1, 5).Font.Bold = $true
    $sheet.Cells.Item(1, 5).Font.Size = 12
    $sheet.Cells.Item(1, 5).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::Tan)

  #Populate the Excel sheet with data from the text file    
for ($i = 0; $i -lt $textContext.Length; $i++) {
    $sheet.cells.Item($i + 2, 1).Value = $textContext[$i]

    $columns = $textContext[$i].Split(' ', [StringSplitOptions]::RemoveEmptyEntries)
    for ($j = 0; $j -lt $columns.Length; $j++) {
        $sheet.cells.item($i + 2, $j + 1).Value = $columns[$j]
    }
}

#Autofit and Font Name the columns
$sheet.UsedRange.Font.Name = "Times New Roman"
$sheet.UsedRange.Font.Size = 12
$sheet.UsedRange.Cells.EntireColumn.AutoFit()

#Save the workbook
$workbook.SaveAs($outputTxt)

#Clean up
$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

}
commandlr