param(
    [string]$SmtpServer = "smtp.gmail.com",
    [int]$SmtpPort = 587,
    [string]$FromEmail,
    [string]$ToEmail,
    [switch]$SendEmail
)

Write-Host "Starting Script..." -ForegroundColor Green

#get network conections and filter for listen
$connections = Get-NetTCPConnection | Where-Object { $_.State -eq 'Listen' -or $_.State -eq 'Established' }

#arrays for each port range
$ports0to1000 = @()
$ports1000to10000 = @()
$ports10000to20000 = @()
$ports20000AndAbove = @()

#get process info
foreach ($conn in $connections) {
    $pId = $conn.OwningProcess
    $localPort = $conn.LocalPort
    
    try {
        $process = Get-Process -Id $pId -ErrorAction SilentlyContinue
        $processName = if ($process) { $process.ProcessName } else { "Unknown" }
        
        $entry = [PSCustomObject]@{
            ProcessName = $processName
            PID = $pId
        }
        
        #categorize the port range
        if ($localPort -le 1000) {
            $ports0to1000 += $entry
        }
        elseif ($localPort -le 10000) {
            $ports1000to10000 += $entry
        }
        elseif ($localPort -le 20000) {
            $ports10000to20000 += $entry
        }
        else {
            $ports20000AndAbove += $entry
        }
    }
    catch {
        Write-Warning "Could not get process info for PID: $pId"
    }
}

#remove duplications
$ports0to1000 = $ports0to1000 | Sort-Object ProcessName, PID -Unique
$ports1000to10000 = $ports1000to10000 | Sort-Object ProcessName, PID -Unique
$ports10000to20000 = $ports10000to20000 | Sort-Object ProcessName, PID -Unique
$ports20000AndAbove = $ports20000AndAbove | Sort-Object ProcessName, PID -Unique

#save results into csv files
$ports0to1000 | Export-Csv -Path "1000.csv" -NoTypeInformation
$ports1000to10000 | Export-Csv -Path "10000.csv" -NoTypeInformation
$ports10000to20000 | Export-Csv -Path "20000.csv" -NoTypeInformation
$ports20000AndAbove | Export-Csv -Path "20000AndAbove.csv" -NoTypeInformation

Write-Host "CSV files created successfully:" -ForegroundColor Green
Write-Host "- 1000.csv: $($ports0to1000.Count) entries"
Write-Host "- 10000.csv: $($ports1000to10000.Count) entries"
Write-Host "- 20000.csv: $($ports10000to20000.Count) entries"
Write-Host "- 20000AndAbove.csv: $($ports20000AndAbove.Count) entries"

#email functionality
if ($SendEmail) {
    if (-not $FromEmail) { $FromEmail = Read-Host "Enter sender email" }
    if (-not $ToEmail) { $ToEmail = Read-Host "Enter recipient email" }
    $securePassword = Read-Host "Enter email password" -AsSecureString
    
    try {
        $credential = New-Object System.Management.Automation.PSCredential($FromEmail, $securePassword)
        
        $mailParams = @{
            SmtpServer = $SmtpServer
            Port = $SmtpPort
            UseSsl = $true
            Credential = $credential
            From = $FromEmail
            To = $ToEmail
            Subject = "Process Port Analysis Report - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            Body = "Please find attached the process port analysis CSV files.`n`nSummary:`n- Ports 0-1000: $($ports0to1000.Count) processes`n- Ports 1001-10000: $($ports1000to10000.Count) processes`n- Ports 10001-20000: $($ports10000to20000.Count) processes`n- Ports 20001+: $($ports20000AndAbove.Count) processes"
            Attachments = @("1000.csv", "10000.csv", "20000.csv", "20000AndAbove.csv")
        }
        
        Send-MailMessage @mailParams
        Write-Host "Email sent successfully!" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to send email: $($_.Exception.Message)"
    }
} else {
    Write-Host "Email sending skipped.use -SendEmail switch to enable email functionali." -ForegroundColor Yellow
}

Write-Host "Script completed!!!!!!!" -ForegroundColor Green
