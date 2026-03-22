#################################################################
# IN WORKSPACE PARAMS IT WILL HAVE THE SERVER URL FOR CONNECTION
#################################################################

$validChoices = @("Commercial", "GCC", "GCC High", "DOD")
do {
    Write-Host "Choose Impact Level: Commercial, GCC, GCC High, DOD"
    $ILChoice = Read-Host "Selection"
} while ($ILChoice -notin $validChoices)

$API = switch ($ILChoice) {
    "Commercial" { "api.powerbi.com" }
    "GCC"        { "api.powerbigov.us" }
    "GCC High"   { "api.high.powerbigov.us" }
    "DOD"        { "api.mil.powerbigov.us" }
}
Write-Host "Using API endpoint: $API"


$Tenant     = Read-Host "Tenant Name"
$json       = Get-Content "C:\Users\${Tenant}WorkspaceParams.json" -ErrorAction Stop | ConvertFrom-Json  #Input JSON File with Workspace Parameters here
$ReportName = "TestFile.pbix" #Input File Name Here
$FilePath   = "C:\Users\$ReportName" #Input File Path Here

#################################################################
# ASSESS IF POWER BI MODULE IS INSTALLED
#################################################################

if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    try {
        Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "Power BI Module installed successfully."
    }
    catch {
        Write-Error "Power BI Module installation failed: $($_.Exception.Message)"
        exit 1
    }
}
else {
    Write-Host "Power BI Module already installed. Continuing script."
}

#################################################################
# CONNECT TO POWER BI
#################################################################
$Environment = switch ($ILChoice) {
    "Commercial" { "Public" }
    "GCC"        { "USGov" }
    "GCC High"   { "USGovHigh" }
    "DOD"        { "USGovMil" }
}

# --- Service Principal (optional) ---
# $TenantId  = Read-Host "Tenant ID"
# $AppId     = Read-Host "App ID"
# $AppSecret = Read-Host "Secret"
# $Credential = New-Object PSCredential($AppId, (ConvertTo-SecureString -String $AppSecret -AsPlainText -Force))
# Connect-PowerBIServiceAccount -$Environment -Tenant $TenantId -ServicePrincipal -Credential $Credential

try {
    Connect-PowerBIServiceAccount -Environment $Environment -ErrorAction Stop
    Write-Host "Connected to Power BI ($ILChoice) successfully."
}
catch {
    Write-Error "Failed to connect to Power BI: $($_.Exception.Message)"
    exit 1
}

#################################################################
# LOOP THROUGH WORKSPACES AND DEPLOY REPORT
#################################################################

foreach ($WorkspaceName in $json.PSObject.Properties.Name) {
    Write-Host "`n--- Processing workspace: $WorkspaceName ---"

    try {
        $Environment = $json.$WorkspaceName.updateDetails[0].URL
        Write-Host "URL: $Environment"

        $WorkspaceId = (Get-PowerBIWorkspace -Name $WorkspaceName -ErrorAction Stop).Id


        # Upload / overwrite report
        Write-Host "Uploading '$ReportName' to workspace '$WorkspaceName'..."
        New-PowerBIReport -Path $FilePath -WorkspaceId $WorkspaceId -ConflictAction Overwrite -ErrorAction Stop
        Write-Host "Report uploaded."

        # Brief wait for dataset to register then look it up by name
        Start-Sleep -Seconds 5
        $ReportBaseName = [System.IO.Path]::GetFileNameWithoutExtension($ReportName)
        $DatasetId = (Get-PowerBIDataset -WorkspaceId $WorkspaceId | Where-Object { $_.Name -eq $ReportBaseName }).Id

        if (-not $DatasetId) {
            Write-Warning "Could not find dataset '$ReportBaseName' in '$WorkspaceName'. Skipping parameter update."
            continue
        }

        Write-Host "Found DatasetId: $DatasetId"

        # Build parameter update body
        $Body = @{
            updateDetails = @(
                @{
                    name     = "Environment"
                    newValue = $Environment
                }
            )
        } | ConvertTo-Json -Depth 10
        
        $Url = "https://$API/v1.0/myorg/groups/$WorkspaceId/datasets/$($DatasetId)/Default.UpdateParameters"
        Invoke-PowerBIRestMethod -Method POST -Url $Url -Body $Body -ContentType "application/json" -ErrorAction Stop

        Write-Host "Parameters updated successfully for '$WorkspaceName'."
    }
    catch {
        Write-Warning "Error processing workspace '$WorkspaceName': $($_.Exception.Message)"
        continue
    }
}

Write-Host "`nScript complete."
Disconnect-PowerBIServiceAccount