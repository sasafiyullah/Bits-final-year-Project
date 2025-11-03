# =================================================================
# SPN Credential Expiry Report and Notification Script
# =================================================================

# --- Login and Graph Connect ---
Connect-AzAccount -Identity
Connect-MgGraph -Identity

# --- Configuration ---
$storageAccountName = "spnreporting"
$resourceGroupName = "Automated-Expiry-Notification"
$containerName = "spnreports"
$blobName = "spn-data.csv"
$tempCsvPath = "$env:TEMP\$blobName"
$date = Get-Date
$alertDays = @(1,2,3,4,5,6,7,15,30,89)

# Retrieve SendGrid API Key securely from Azure Automation Variable
Write-Host "Retrieving SendGrid API Key from Automation Variables..."
$SendGridApiKey = Get-AutomationVariable -Name "SendGridApiKey"
if (-not $SendGridApiKey) {
    Write-Error "Fatal Error: Could not retrieve the 'SendGridApiKey' Automation Variable. Check variable existence and permissions. Exiting script."
    # Disconnect before exiting for cleanup
    Disconnect-AzAccount -Confirm:$false
    Disconnect-MgGraph
    Exit 1
}

$FromEmail = "sasafiyullah@outlook.com"
$FromName = "SPN-Notification"

# --- Storage Context using Storage Key ---
try {
    $storageKey = (Get-AzStorageAccountKey -ResourceGroupName $resourceGroupName -Name $storageAccountName)[0].Value
    $ctx = New-AzStorageContext -StorageAccountName $storageAccountName -StorageAccountKey $storageKey
} catch {
    Write-Error "Failed to create Azure Storage context. Check resource group/storage account name or permissions."
    Exit 1
}

# --- Cleanup Old Reports ---
write-host "--- Cleaning up old local and blob reports... ---"
# Clean Old CSVs in TEMP
Get-ChildItem -Path $env:TEMP | Where-Object { $_.Name -like "spn-data*" } | Remove-Item -Force -ErrorAction SilentlyContinue
# Clean Old Blobs
Get-AzStorageBlob -Container $containerName -Context $ctx | Where-Object { $_.Name -like "spn-data*" } | ForEach-Object {
    Remove-AzStorageBlob -Container $containerName -Blob $_.Name -Context $ctx
}
write-host "Successfully deleted the old reports."

# --- Fetch SPNs and Owners ---
write-host "--- Fetching all applications and credentials from Microsoft Graph... ---"
$spns = Get-MgApplication -All
$results = @()

foreach ($spn in $spns) {
    try {
        # FIX: Get full application details (including credentials) in one efficient call
        $fullApplication = Get-MgApplication -ApplicationId $spn.Id -Property "DisplayName,KeyCredentials,PasswordCredentials"
        $owners = Get-MgApplicationOwner -ApplicationId $spn.Id
        
        $ownerNames = $owners.AdditionalProperties.displayName -join ";"

        $emailList = @()
        foreach ($owner in $owners) {
            $email = $owner.AdditionalProperties.mail
            if (-not [string]::IsNullOrEmpty($email)) {
                if ($email -match "_") {
                    $emailList += $email.Split("_")[1]
                } else {
                    $emailList += $email
                }
            }
        }
        $ownerEmails = $emailList -join ","

        $creds = $fullApplication.KeyCredentials
        $secrets = $fullApplication.PasswordCredentials
        $allCreds = $creds + $secrets

        foreach ($cred in $allCreds) {
            if ($cred -is [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]) {
                $type = "Certificate"
            } else {
                $type = "Secret"
            }

            $results += [PSCustomObject]@{
                Name        = $spn.DisplayName
                ExpiryDate  = $cred.EndDateTime.ToString("yyyy-MM-dd")
                Type        = $type
                OwnerName   = $ownerNames
                OwnerEmail  = $ownerEmails
            }
        }
    } catch {
        Write-Warning "Could not process application $($spn.DisplayName) ($($spn.Id)). Error: $($_.Exception.Message)"
    }
}

# --- Save to CSV and Upload to Blob ---
write-host "--- Saving data to CSV and uploading to blob storage... ---"
$results | Export-Csv -Path $tempCsvPath -NoTypeInformation
Set-AzStorageBlobContent -File $tempCsvPath -Container $containerName -Blob $blobName -Context $ctx -Force
write-host "Report saved to $containerName/$blobName."


# =================================================================
# Email Functions and Styling
# =================================================================

# --- HTML Styling ---
$HtmlHead = @'
<style>
    body { font-family: Calibri, sans-serif; background-color: #f4f4f4; padding: 20px; }
    .container { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); }
    table { border-collapse: collapse; width: 100%; margin-top: 15px; }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #facb48; font-weight: bold; }
    .alert-header { color: #cc0000; font-size: 1.2em; margin-bottom: 10px; }
</style>
'@

function Get-HtmlTable {
    param ([array]$Rows)
    $rowsHtml = $Rows | ForEach-Object {
        "<tr><td>$($_.SPN)</td><td>$($_.Expiry)</td><td>$($_.Type)</td><td>$($_.Owner)</td><td>$($_.Email)</td></tr>"
    }
    return @"
<table>
<tr><th>SPN</th><th>Expiry</th><th>Type</th><th>Owner</th><th>Email</th></tr>
$rowsHtml
</table>
"@
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
        from = @{ email = $FromEmail; name = $FromName }
        content = @(@{ type = "text/html"; value = $Body })
    }

    try {
        $jsonBody = $payload | ConvertTo-Json -Depth 10
        Invoke-RestMethod -Uri "https://api.sendgrid.com/v3/mail/send" -Method POST `
            -Headers @{ Authorization = "Bearer $SendGridApiKey" } -ContentType "application/json" `
            -Body $jsonBody -MaximumRedirection 0 # Added to prevent unnecessary redirects
        
        # --- ENHANCEMENT: SUCCESS LOG MESSAGE ---
        Write-Host "Successfully sent email: '$Subject' to $($validToEmails -join ', ')"
        
    } catch {
        # --- ERROR LOG MESSAGE ---
        Write-Error "Failed to send email via SendGrid: $($_.Exception.Message)"
    }
}


# =================================================================
# Alerting Logic
# =================================================================

write-host "--- Checking for expiring SPNs and sending alerts... ---"

# Retrieve the latest CSV from the blob
Get-AzStorageBlobContent -Container $containerName -Blob $blobName -Destination $tempCsvPath -Context $ctx -Force
$spnData = Import-Csv $tempCsvPath

foreach ($spn in $spnData) {
    $expiryDate = [datetime]$spn.ExpiryDate
    $daysLeft = ($expiryDate - $date).Days

    if ($daysLeft -in $alertDays) {
        $row = [PSCustomObject]@{
            SPN     = $spn.Name
            Expiry  = $expiryDate.ToString("dd-MMM-yyyy")
            Type    = $spn.Type
            Owner   = $spn.OwnerName
            Email   = $spn.OwnerEmail
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
<p>Regard</p>
<p>Cloud-Admin</p>
</div>
</body>
</html>
"@
        # Split and clean recipient emails
        $recipients = $row.Email.Split(",", [System.StringSplitOptions]::RemoveEmptyEntries) | ForEach-Object { $_.Trim() }

        Send-EmailViaSendGrid `
            -FromEmail $FromEmail `
            -FromName $FromName `
            -ToEmails $recipients `
            -Subject "SPN $($row.Type) Expiry Alert: $($row.SPN) - $daysLeft Days Remaining" `
            -Body $htmlBody `
            -SendGridApiKey $SendGridApiKey
    }
}

# --- Disconnect Sessions (Best Practice) ---
write-host "--- Disconnecting sessions. ---"
Disconnect-AzAccount -Confirm:$false
Disconnect-MgGraph
write-host "Script execution complete."