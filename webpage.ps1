$url = "https://www.microsoft.com"
$response = Invoke-WebRequest -Uri $url

# Extract the title of the page
$pageTitle = $response.ParsedHtml.title

# Find all links on the page
$links = $response.Links | Select-Object -ExpandProperty href

Write-Host "Page Title: $pageTitle" -ForegroundColor Yellow 
Write-Host "Links on the page:"     -ForegroundColor Green
$links | ForEach-Object { Write-Host $_ }

$links | Out-File -FilePath c:\mstest.txt