# Get a list of all .html files in the current directory

$Dir = Get-ChildItem .
$List = $Dir | where {$_.Extension -eq ".html"}
 Write-Host "Available templates: $List" -ForegroundColor Green

 $HTMLFile = $( Read-Host "Filename (include '.html')" )
 $Subject = $( Read-Host "Email subject: " )
 $Recipient = $( Read-Host "Who do you want to send the email to? :" )
# Check if the personalised email template exists before continuing.

Try {
    $InputFile = "./$HTMLFile"
    $HTML = Get-Content -Path $InputFile -Raw -ErrorAction Stop
} Catch {
    Write-Host "ERROR: The email template could not be found." -ForegroundColor Red
    Break
}

# This fails on PowerShell Mac, so guessing the error is close enough.

Try {
    $Outlook = New-Object -com Outlook.Application
} Catch {
    Write-Host "ERROR: Cannot create a new Outlook object. Outlook may not be installed on your machine." -ForegroundColor Red
    Break
}

$Mail = $outlook.CreateItem(0)
$Mail.to = $Recipient
$Mail.Subject = $Subject
$Mail.HTMLBody =$HTML
Write-Output "Sending email to: $Recipient"
$Mail.Send()
Write-Host "All done :)" -ForegroundColor Green
$outlook.Quit()
