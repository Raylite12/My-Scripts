#Define the path for the output Excel file
$outputCsv = "c:\test"

#Create an Excel COM object
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True      #Make Excel visible (remove or set to $false for background operation)

#Create a new workbook
$workbook = $Excel.Workbooks.Add()

#Get the first worksheet
#Access the first default sheet and rename it
$sheet = $workbook.Sheets.item(1)   #Access the first sheet by index
$sheet.Name = "Compound Int"

#Add a title to the worksheet
$Sheet.Cells.Item(1, 1).Value          = "Starting Amount"
$Sheet.Cells.Item(1, 1).Font.Bold      = $True
$Sheet.Cells.Item(1, 1).Font.Size      = 12       #You could delete this line if you use ($sheet.usedRange.Font.Size = 12 This will change the font for the whole Excel)
$Sheet.Cells.Item(1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(2, 1).Value          = "Intrest Rate"
$Sheet.Cells.Item(2, 1).Font.Bold      = $True
$Sheet.Cells.Item(2, 1).Font.Size      = 12       
$Sheet.Cells.Item(2, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(3, 1).Value          = "Times Per Year"
$Sheet.Cells.Item(3, 1).Font.Bold      = $True
$Sheet.Cells.Item(3, 1).Font.Size      = 12       
$Sheet.Cells.Item(3, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(4, 1).Value          = "Years"
$Sheet.Cells.Item(4, 1).Font.Bold      = $True
$Sheet.Cells.Item(4, 1).Font.Size      = 12       
$Sheet.Cells.Item(4, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(5, 1).Value          = "Monthly Deposit"
$Sheet.Cells.Item(5, 1).Font.Bold      = $True
$Sheet.Cells.Item(5, 1).Font.Size      = 12       
$Sheet.Cells.Item(5, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(6, 1).Value          = "Final Amount"
$Sheet.Cells.Item(6, 1).Font.Bold      = $True
$Sheet.Cells.Item(6, 1).Font.Size      = 12       
$Sheet.Cells.Item(6, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(8, 1).Value          = "Years"
$Sheet.Cells.Item(8, 1).Font.Bold      = $True
$Sheet.Cells.Item(8, 1).Font.Size      = 12       
$Sheet.Cells.Item(8, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$Sheet.Cells.Item(8, 2).Value          = "Amount"
$Sheet.Cells.Item(8, 2).Font.Bold      = $True
$Sheet.Cells.Item(8, 2).Font.Size      = 12       
$Sheet.Cells.Item(8, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::LightGreen)

$range = $sheet.Range("A9:A39")
$range.Select()

$rows = 32
$data = New-Object 'object[,]' $rows, 1
for ($i = 0; $i -lt $rows; $i++) {
    $data[$i,0] = $i
}

$Sheet.Range("A9:A39").Value2 = $data

#Add info to the worksheet
$Sheet.Cells.Item(2, 2).Value = "7%"
$Sheet.Cells.Item(3, 2).Value = "12"
$Sheet.Cells.Item(4, 2).Value = "5"
$Sheet.Cells.Item(6, 2).Value = "=B1*(1+B2/B3)^(B3*B4)+B5*((1+B2/B3)^(B3*B4)-1)/(B2/B3)"
$Sheet.Cells.Item(9, 2).Value = '=$B$1*(1+$B$2/$B$3)^($B$3*A9)+$B$5*((1+$B$2/$B$3)^($B$3*A9)-1)/($B$2/$B$3)'
#Fill down B9
$Sheet.Range("B9:B39").FillDown()

#Populate the Excel sheet with data from the text file
for ($i = 0; $i -lt $textContent.Length; $i++) {
    $Sheet.Cells.Item($i + 2, 1).Value = $textContent[$i]

    $columns = $textContent[$i].Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
    for ($j = 0; $j -lt $columns.Length; $j++) {
        $Sheet.Cells.Item($i + 2, $J + 1).Value = $columns[$j]
    }
}

#$Sheet.Cells["B2:B$($Sheet.Dimension.End.Row)] | Set-format -HorizontalAlignment Right

#Autofit and Font Name the columns
$Sheet.UsedRange.Font.Name = "Times New Roman"
$Sheet.UsedRange.Font.Size = 12
#$Sheet.UsedRange.HorizontalAlignment = -4152    #To line the whole Excel to the Align Right
$Sheet.UsedRange.Cells.EntireColumn.AutoFit()

#Save the workbook
$workbook.SaveAs($outputCsv)

Write-Host "Data exported to CSV: $outputCsv"          -ForegroundColor Cyan
Write-Host "You can now open this CSV file in Excel"   -ForegroundColor Green

#Clean up 
#$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Sheet)       | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)    | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)       | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
