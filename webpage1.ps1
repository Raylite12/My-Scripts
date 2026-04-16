#Phase 1: Retrieve and Parse Data
#Define the URL of the target website
$url = "https://learn.microsoft.com/en-us/docs/"

#Start
Start-Process $url

#Send a GET request to the URL
$response = Invoke-WebRequest -Uri $url

#Access the parsed HTML content (DOM)
$html = $response.ParsedHtml

#Extract a single element by ID
$data = $html.getElementById("dataID").innerText

#Extract a table of data
$tables = $html.getElementsByTagName("table")

#You might need to iterate through $tables to find the correct one
#For example, selecting the first table
$targetTable = $tables[0]
#Process rows and cells as needed

#Extract all links 
$links = $response.Links.Href

#Phase 2: Output to Excel/CSV
#Structure your data into PowerShell objects
$dataObject = New-Object PSObject -Property @{
    Column1 = $data
    #Add more properties as needed 
}

#Initialize dataArray if not already initialized
if (-not $dataArray) {
   $dataArray = @()
}

#For multiple rows of data, add objects to an array
$dataArray += $dataObject

#Export the object(s) to a CSV file
$csvPath = "c:test"
$dataArray | Export-Csv -Path $csvPath -NoTypeInformation

#Phase 3: Input into Another Website
#Submit the data
$formUrl = "www.google.com"
$formBody = @{
    inputField1 = "value1"
    inputField2 = "value2"
    #add all required fields and their values
}
Invoke-WebRequest -Uri $formUrl -Method Post -body $formBody -SessionVariable session