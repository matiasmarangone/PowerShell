import-module ActiveDirectory

$ExportPath = "C:\Accounts\users_password_expiring_date.csv"
$ExportBackupPath = "C:\Accounts\Backup\users_password_expiring_date " + (Get-Date -Format "yyyy-MM-dd") + ".csv"

$filterDateMax = (Get-Date).AddDays(14).Date
$filterDateMin = (Get-Date).Date

$users = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False -and DisplayName -like "*" -and  ( (Description -like "*John Doe*") -or (Description -like "*John Doe*") )} –Properties "DisplayName", "sAMAccountName", "mail", "description", "msDS-UserPasswordExpiryTimeComputed", "accountExpirationDate" | Select-Object -Property "Displayname","sAMAccountName", "mail", "description", @{Name="PasswordExpiryDate";Expression={[datetime]::FromFileTime($PSItem."msDS-UserPasswordExpiryTimeComputed")}}, "accountExpirationDate" | Where-Object {($PSItem.PasswordExpiryDate -lt $filterDateMax ) -and ($PSItem.PasswordExpiryDate -gt $filterDateMin ) }

$users = $users | Select-Object *, @{Name="DaysToExpire";Expression={(New-TimeSpan -Start $filterDateMin  -End $PSItem.PasswordExpiryDate ).Days.ToString()}} | Sort-Object -Property PasswordExpiryDate -Descending

$users | Select-Object * | Export-csv -Path $ExportPath

Copy-Item $ExportPath -Destination $ExportBackupPath

##Send emails to users
$From = "*"
$Cc = "*"
$Subject = "Email subject"
$Attachement = "C:\Accounts\resources\example.pdf"
$SMTPServer = "*"
$SMTPPort = "*"

foreach ($user in $users) {

 $To = $user.mail

 $Body = "Hello, you are the user (" +  ($user.sAMAccountName) +  "), your account will expire in " + ($user.DaysToExpire).ToString() +" days"

 ##Write-Host "Sending Email ", ($user.mail).ToString()

 Send-MailMessage -From $From -to $To -Cc $Cc -Subject $Subject -Body $Body  -Attachments $Attachement -SmtpServer $SMTPServer -port $SMTPPort  -Encoding UTF8


 }
