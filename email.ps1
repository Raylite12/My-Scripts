$email = Read-Host "Enter Email"

$Message = new-object Net.Mail.MailMessage
$smtp = new-object Net.Mail.smtpClient("smtp.yahoo.com")
#smtp.EnableSss1 = $true

$smtp.Timeout = 400000
$Message.From = "test@yahoo.com"
$Message.To.Add($email)
$Message.Subject = "Put in a Subject"
$Message.Body = ""  #if you want to get from file: Get-Content -Path "path/to/text" -Encoding UTF8 -Raw
#$Message.Attachment.Add("file/path")
$smtp.Send($Message)