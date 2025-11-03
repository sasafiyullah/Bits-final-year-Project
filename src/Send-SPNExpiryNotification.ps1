# =================================================================
# SPN Credential Expiry Report and Notification Script
# =================================================================

# --- Login and Graph Connect ---
Connect-AzAccount -Identity
Connect-MgGraph -Identity -NoWelcome

# --- Configuration ---
$storageAccountName = "spnreporting"
$resourceGroupName  = "Automated-Expiry-Notification"
$containerName      = "spnreports"
$date               = Get-Date

# === Date-stamped filenames ===
$todayString  = $date.ToString("yyyy-MM-dd")
$blobName     = "$todayString-spn-data.csv"
$tempCsvPath  = Join-Path $env:TEMP $blobName
$retentionDays = 10
$alertDays     = @(1,2,3,4,5,6,7,15,30,89,90)

# --- Retrieve SendGrid API Key securely from Azure Automation Variable ---
Write-Host "Retrieving SendGrid API Key from Automation Variables..."
$SendGridApiKey = Get-AutomationVariable -Name "SendGridApiKey"
if (-not $SendGridApiKey) {
    Write-Error "Fatal Error: Could not retrieve the 'SendGridApiKey' Automation Variable."
    Disconnect-AzAccount -Confirm:$false
    Disconnect-MgGraph
    Exit 1
}

$FromEmail = "sasafiyullah@outlook.com"
$FromName  = "SPN-Notification"

# --- Storage Context ---
try {
    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
    $ctx        = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey
} catch {
    Write-Error "Failed to create Azure Storage context. Error: $($_.Exception.Message)"
    Disconnect-AzAccount -Confirm:$false
    Disconnect-MgGraph
    Exit 1
}

# === Cleanup: 30-day retention policy ===
Write-Host "--- Cleaning up old local and blob reports... ---"
Get-ChildItem -Path $env:TEMP -Filter "*-spn-data.csv" -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

$retentionThreshold = (Get-Date).AddDays(-$retentionDays)
$deletedCount = 0
Get-AzStorageBlob -Container $containerName -Context $ctx -ErrorAction SilentlyContinue | Where-Object {
    $_.Name -like "*-spn-data.csv" -and $_.LastModified.UtcDateTime -lt $retentionThreshold
} | ForEach-Object {
    Remove-AzStorageBlob -Container $containerName -Blob $_.Name -Context $ctx -ErrorAction SilentlyContinue
    $deletedCount++
}
Write-Host "Deleted $deletedCount old report(s)."

# =================================================================
# HTML Styling & Email
# =================================================================

$HtmlHead = @'
<style>
    body { font-family: Calibri, sans-serif; background-color: #f4f4f4; padding: 20px; }
    .container { background-color: white; padding: 20px; border-radius: 8px; }
    table { border-collapse: collapse; width: 100%; margin-top: 15px; }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #facb48; font-weight: bold; }
    .alert-header { color: #cc0000; font-size: 1.2em; margin-bottom: 10px; }
    .footer { color: #666; font-size: 0.85em; margin-top: 15px; }
</style>
'@

function Get-HtmlTable {
    param ([array]$Rows)
    $rowsHtml = $Rows | ForEach-Object {
        "<tr><td>$($_.SPN)</td><td>$($_.Expiry)</td><td>$($_.Type)</td><td>$($_.Owner)</td><td>$($_.Email)</td></tr>"
    }
    $table = @"
<table>
<tr><th>SPN</th><th>Expiry</th><th>Type</th><th>Owner</th><th>Email</th></tr>
$rowsHtml
</table>
"@
    return $table
}

function Send-EmailViaSendGrid {
    param (
        [string]$FromEmail,
        [string]$FromName,
        [string[]]$ToEmails,
        [string]$Subject,
        [string]$Body,
        [string]$SendGridApiKey
    )

    $validToEmails = $ToEmails | Where-Object { -not [string]::IsNullOrEmpty($_) }
    if (-not $validToEmails) {
        Write-Warning "No valid recipients found for email: $Subject. Skipping."
        return
    }

    $toRecipients = @()
    foreach ($email in $validToEmails) {
        $toRecipients += @{ email = $email }
    }

    $payload = @{
        personalizations = @(@{ to = $toRecipients; subject = $Subject })
        from             = @{ email = $FromEmail; name = $FromName }
        content          = @(@{ type = "text/html"; value = $Body })
    }

    try {
        $jsonBody = $payload | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" -Method POST `
            -Headers @{
                Authorization = "Bearer $SendGridApiKey"
                "Content-Type" = "application/json"
            } `
            -Body $jsonBody -MaximumRedirection 0
        Write-Host "Successfully sent email: '$Subject' to $($validToEmails -join ', ')"
    } catch {
        Write-Error "Failed to send email via SendGrid: $($_.Exception.Message)"
    }
}

# =================================================================
# Fetch SPNs and Owners
# =================================================================

Write-Host "--- Fetching all applications and credentials from Microsoft Graph... ---"

$spns     = Get-MgApplication -All -Property "Id,DisplayName,KeyCredentials,PasswordCredentials"
$results  = @()

foreach ($spn in $spns) {
    try {
        $fullApplication = Get-MgApplication -ApplicationId $spn.Id -Property "DisplayName,KeyCredentials,PasswordCredentials"
        $owners = Get-MgApplicationOwner -ApplicationId $spn.Id -ErrorAction SilentlyContinue

        $ownerNames  = $null
        $ownerEmails = $null

        if ($owners) {
            $ownerNames = ($owners | ForEach-Object { $_.AdditionalProperties.displayName }) -join ";"

            $emailList = @()
            foreach ($owner in $owners) {
                $ap = $owner.AdditionalProperties
                # Prefer user.mail (guests and members), include group mail if present; skip SPNs
                if ($ap.'@odata.type' -eq "#microsoft.graph.user") {
                    if ($ap.mail) { $emailList += $ap.mail }
                } elseif ($ap.'@odata.type' -eq "#microsoft.graph.group") {
                    if ($ap.mail) { $emailList += $ap.mail }
                }
            }
            $ownerEmails = ($emailList | Where-Object { $_ }) -join ","
        }

        $creds    = $fullApplication.KeyCredentials
        $secrets  = $fullApplication.PasswordCredentials
        $allCreds = @()
        if ($creds)   { $allCreds += $creds }
        if ($secrets) { $allCreds += $secrets }

        foreach ($cred in $allCreds) {
            if (-not $cred.EndDateTime) { continue }
            $type = if ($cred -is [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]) { "Certificate" } else { "Secret" }

            $results += [PSCustomObject]@{
                Name       = $fullApplication.DisplayName
                ExpiryDate = $cred.EndDateTime.ToString("yyyy-MM-dd")
                Type       = $type
                OwnerName  = $ownerNames
                OwnerEmail = $ownerEmails
            }
        }
    } catch {
        Write-Warning "Could not process application $($spn.DisplayName) ($($spn.Id)). Error: $($_.Exception.Message)"
    }
}

# =================================================================
# Save to CSV and Upload
# =================================================================

Write-Host "--- Saving data to CSV and uploading to blob storage... ---"
if ($results -and $results.Count -gt 0) {
    $results | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8
    Set-AzStorageBlobContent -File $tempCsvPath -Container $containerName -Blob $blobName -Context $ctx -Force | Out-Null
    Write-Host "Report saved to $containerName/$blobName."
} else {
    Write-Host "No results found. Skipping CSV export and upload."
}

# =================================================================
# Alerting Logic
# =================================================================

Write-Host "--- Checking for expiring SPNs and sending alerts... ---"

if (-not $results -or $results.Count -eq 0) {
    Write-Host "No SPN data found or processed. Skipping alerting."
    Disconnect-AzAccount -Confirm:$false
    Disconnect-MgGraph
    Write-Host "Script execution complete."
    Exit 0
}

foreach ($spn in $results) {
    $expiryDate = [datetime]$spn.ExpiryDate
    $daysLeft   = ($expiryDate.Date - $date.Date).Days

    if ($daysLeft -in $alertDays) {
        $row = [PSCustomObject]@{
            SPN    = $spn.Name
            Expiry = $expiryDate.ToString("dd-MMM-yyyy")
            Type   = $spn.Type
            Owner  = $spn.OwnerName
            Email  = $spn.OwnerEmail
        }

        $tableHtml = Get-HtmlTable @($row)
        $htmlBody = @"
<!DOCTYPE html>
<html>
<head>$HtmlHead</head>
<body>
<div class="container">
<p>Hello,</p>
<p class="alert-header">Your $($row.Type) for SPN <b>$($row.SPN)</b> is expiring in <b>$daysLeft</b> days on $($row.Expiry).</p>
$tableHtml
<p>Please renew this credential immediately to avoid service disruption.</p>
<p>Regards,</p>
<p>Cloud-Admin</p>
<p class="footer">This is an automated message from SPN-Expiry-Automation. Do not reply to this email.</p>
</div>
</body>
</html>
"@

        $recipients = @()
        if ($row.Email) {
            $recipients = $row.Email.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }
        }
        if (-not $recipients -or $recipients.Count -eq 0) {
            Write-Warning "No owner emails found for '$($row.SPN)'. Skipping alert email."
            continue
        }

        Send-EmailViaSendGrid `
            -FromEmail $FromEmail `
            -FromName  $FromName `
            -ToEmails  $recipients `
            -Subject   "SPN $($row.Type) Expiry Alert: $($row.SPN) - $daysLeft Days Remaining" `
            -Body      $htmlBody `
            -SendGridApiKey $SendGridApiKey
    }
}

# --- Disconnect Sessions (Best Practice) ---
Write-Host "--- Disconnecting sessions. ---"
Disconnect-AzAccount -Confirm:$false
Disconnect-MgGraph
Write-Host "Script execution complete."