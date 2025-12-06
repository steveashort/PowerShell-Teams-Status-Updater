<#
.SYNOPSIS
    Sends a formatted status update to Microsoft Teams from a JSON file.

.DESCRIPTION
    This script reads a key-value JSON file, formats it into a Microsoft Teams 
    MessageCard with a "facts" section, and sends it to a configured Webhook URL.

.PARAMETER WebhookUrl
    The Microsoft Teams Webhook URL. Can be passed as a parameter or set via 
    the TEAMS_WEBHOOK_URL environment variable.

.PARAMETER JsonFilePath
    Path to the JSON data file. Defaults to ".\data.json".

.EXAMPLE
    .\send_update.ps1 -WebhookUrl "https://outlook.office.com/webhook/..."
    
.EXAMPLE
    $env:TEAMS_WEBHOOK_URL = "https://outlook.office.com/webhook/..."
    .\send_update.ps1 -JsonFilePath "C:\Temp\status.json"
#>

param(
    [string]$WebhookUrl = $env:TEAMS_WEBHOOK_URL,
    [string]$JsonFilePath = ".\data.json"
)

# 1. Validation
if ([string]::IsNullOrWhiteSpace($WebhookUrl)) {
    Write-Error "Error: Webhook URL is missing. Please provide it via the -WebhookUrl parameter or TEAMS_WEBHOOK_URL environment variable."
    exit 1
}

if (-not (Test-Path $JsonFilePath)) {
    Write-Error "Error: Data file not found at $JsonFilePath"
    exit 1
}

# 2. Read and Parse JSON Data
try {
    $JsonContent = Get-Content -Path $JsonFilePath -Raw -ErrorAction Stop
    $Data = $JsonContent | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Error "Error reading or parsing JSON file: $_"
    exit 1
}

# 3. Build Facts List for Teams Card
# We iterate over the properties of the PSCustomObject created by ConvertFrom-Json
$Facts = @()

# Handle case where JSON might be an array or single object. Assuming single object based on requirements.
if ($Data -is [array]) {
    $Properties = $Data[0].PSObject.Properties
} else {
    $Properties = $Data.PSObject.Properties
}

foreach ($Prop in $Properties) {
    $Facts += @{
        name  = $Prop.Name
        value = $Prop.Value.ToString()
    }
}

# 4. Construct the MessageCard Payload
$Card = @{
    "@type"      = "MessageCard"
    "@context"   = "http://schema.org/extensions"
    themeColor   = "0076D7"
    summary      = "Status Update"
    sections     = @(
        @{
            activityTitle    = "System Status Update"
            activitySubtitle = "Latest metrics from the desktop client"
            facts            = $Facts
            markdown         = $true
        }
    )
}

# 5. Convert to JSON
$CardJson = $Card | ConvertTo-Json -Depth 10

# 6. Send to Teams
try {
    Write-Host "Sending update to Teams..."
    $Response = Invoke-RestMethod -Uri $WebhookUrl -Method Post -Body $CardJson -ContentType "application/json" -ErrorAction Stop
    
    # Microsoft Teams Webhook usually returns 1 on success
    if ($Response -eq 1 -or $Response -eq "1") {
        Write-Host "Success: Update sent to Teams." -ForegroundColor Green
    }
    else {
        Write-Host "Update sent, response: $Response" -ForegroundColor Cyan
    }
}
catch {
    Write-Error "Failed to send to Teams: $_"
    exit 1
}
