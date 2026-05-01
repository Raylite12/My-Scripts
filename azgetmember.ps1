[net.servicePointManager]::SecurityProtocol = 
[Net.ServicePointManager]::SecurityProtocol -Bor 
[Net.SecurityProtocolType]::Tls12

#Connect to AZ
Connect-AzAccount -Tenant
Get-AzContext | Format-List *
Connect-MgGraph
Get-MgContext | format-list *
Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Groups

#Get user DisplayName and ID 
Get-MgGroupMember -GroupId 000000-0000-0000-00000 |
     Select-Object @{Name="Display"; Expression={$_.AdditionalProperties["display"]}}, Id |
	 Out-File "c:the path"
	 
#Function to generate a random password
function New-RandowPassword {
    param([int)$length = 20)
	
	$characterSet = [char[]](48..57 + 65.90 + 97..122 + 33..47 + 58..64 + 91..96 + 123..126)
	$StringSet = 1..$length | ForEach-Object { Get-Random -InputObject $CharacterSet }
	
	-join $StringSet
}
	 
#Define the path to your input file
#Define the path for the output Excel file
#Get the context of the text file 	 
$input = "c:the path"
$output = "c:the path"
$textContent = Get-Content $input

#Create an Excel Com Object
$Excel = New-Object -ComObject Excel.Appliation
$Excel.Visible = $True        #Make Excel visible (remove or set false for background operation)

#Create a new workbook
    $workbook = $excel.workbooks.Add()

    #Get the first worksheet 
    #Access the first default sheet and rename it
    $sheetGroup = $workbook.Sheets.Item(1)   #Access the first sheet by index 
    $sheetGroup.Name = "Group Members"

    #Add a title to the worksheet
    $sheetGroup.Cells.Item(1, 1).value = "DisplayName"
    $sheetGroup.Cells.Item(1, 1).Font.Bold = $true
    $sheetGroup.Cells.Item(1, 1).Font.Size = 12
    $sheetGroup.Cells.Item(1, 1).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::YellowGreen)

    $sheetGroup.Cells.Item(1, 2).value = "Id"
    $sheetGroup.Cells.Item(1, 2).Font.Bold = $true
    $sheetGroup.Cells.Item(1, 2).Font.Size = 12
    $sheetGroup.Cells.Item(1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BurlyWood)
	
	$sheetGroup.Cells.Item(1, 2).value = "Temporary Password"
    $sheetGroup.Cells.Item(1, 2).Font.Bold = $true
    $sheetGroup.Cells.Item(1, 2).Font.Size = 12
    $sheetGroup.Cells.Item(1, 2).Interior.Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::BurlyWood)
	
#Populate the Excel sheet with data from the text file and add random passwords
for ($i = 0; $i -lt $textContent.Length; $i++) {
    $columns = $textContent[$i].Spilt(' ', [StringSplitOptions]::RemoveEmptyEntries)
	for ($j = 0; $j -lt $columns.Length; $j++) {
	    $sheet.cells.item($i + 2, $j + 1).Value = $columns[$j]
	}
	
	#Generate a random password for the current user 
	$tempPW = New-RandowPassword
	$sheet.Cells.Item($i + 2, 3).Value = $tempPW
}

#Autofit and font Name the columns
$Sheet.UsedRange.Font.Name = "Times New Roman"
$Sheet.UsedRange.Font.Size = 12
$Sheet.UsedRange.Cells.EntireColumn.AutoFit()

#Save the workbook
$workbook.SaveAs($OutPut)

#Clean Up 
$Excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

#Need the GroupId number
Remove-MgGroupMemberByRef -DirectoryObjectId 0000-00000-0000-0000 GroupId 00000-0000-000-0000

#GroupId OTP
New-MgGroupMember -GroupId 0000-0000-0000-0000 -DirectoryObjectId 0000-0000-0000

<#Import group IDs from CSV and remove the user from each
Import-Csv "c:" | ForEach-Object {
    Remove-MgGroupMemberByRef -GroupId $_.GroupObjectID -DirectoryObjectId "Your User ID"
}#>