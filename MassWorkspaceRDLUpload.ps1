#################################################################
# UPLOAD POWER BI PAGINATED (.rdl) REPORT TO MULTIPLE WORKSPACES
#################################################################

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
# VARIABLES
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


$Tenant      = Read-Host "Tenant Name"
$JsonPath    = "C:\Users\${Tenant}WorkspaceParams.json" #Input JSON File Path Here
$RDLFilePath = "C:\Users\Test File.rdl" #Input File Path Here

if (-not (Test-Path $JsonPath)) {
    Write-Error "Workspace params file not found: $JsonPath"
    exit 1
}
if (-not (Test-Path $RDLFilePath)) {
    Write-Error "RDL file not found: $RDLFilePath"
    exit 1
}

$json = Get-Content $JsonPath -ErrorAction Stop | ConvertFrom-Json

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
# $TenantId   = Read-Host "Tenant ID"
# $AppId      = Read-Host "App ID"
# $AppSecret  = Read-Host "Secret"
# $Credential = New-Object PSCredential($AppId, (ConvertTo-SecureString -String $AppSecret -AsPlainText -Force))
# Connect-PowerBIServiceAccount -Environment $Environment -Tenant $TenantId -ServicePrincipal -Credential $Credential -ErrorAction Stop

# --- Interactive login ---
try {
    Connect-PowerBIServiceAccount -Environment $Environment -ErrorAction Stop
    Write-Host "Connected to Power BI ($ILChoice) successfully."
}
catch {
    Write-Error "Failed to connect to Power BI: $($_.Exception.Message)"
    exit 1
}

#################################################################
# FUNCTION: PUBLISH RDL FILE TO POWER BI
#################################################################

function Publish-RDLFile {
    param (
        [string]$RdlFilePath,
        [string]$GroupId,
        [string]$API
    )

    $fileName = [IO.Path]::GetFileName($RdlFilePath)
    $fileNameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($RdlFilePath)
    $boundary = [guid]::NewGuid().ToString()
    $fileBody = Get-Content -Path $RdlFilePath -Encoding UTF8

    $body = @"
---------FormBoundary$boundary
Content-Disposition: form-data; name="$filename"; filename="$filename"
Content-Type: application/rdl

$fileBody 
---------FormBoundary$boundary--

"@

    # Get AccessToken and set it as header.
    $headers = Get-PowerBIAccessToken

    #Lookup if report is in Workspace Already
    $GetReport = Get-PowerBIReport -Name $fileNameNoExt -WorkspaceId $GroupId

    #If report is in, overwrite it, otherwise upload new
    $nameConflict = if ($GetReport) {
    "Overwrite"
    }
    else {
    "Abort"
    }

    #Set URL
    if ($GroupId) {
        $url = "https://$($API)/v1.0/myorg/groups/$GroupId/imports?datasetDisplayName=$fileName&nameConflict=$nameConflict"
    }
    else {
        $url = "https://$($API)/v1.0/myorg/imports?datasetDisplayName=$fileName&nameConflict=$nameConflict"
    }

    # Create import
    $report = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "multipart/form-data"   
    $report.id
}

#################################################################
# LOOP: UPDATE CONNECTION STRING AND PUBLISH PER WORKSPACE
#################################################################

$success = 0
$failed  = 0

foreach ($WorkspaceName in $json.PSObject.Properties.Name) {
    Write-Host "`n--- Processing workspace: $WorkspaceName ---"

    try {
        $details = $json.$WorkspaceName.updateDetails[0]
        $url     = if ($details.URL) { $details.URL } else { $details.newValue }

        if (-not $url) {
            Write-Warning "No URL found for '$WorkspaceName'. Skipping."
            $failed++
            continue
        }

        # Derive database name from URL (e.g. tmt-demo.crm.dynamics.com -> tmt-demo)
        $database = $url -replace "\..*", ""
        Write-Host "URL: $url"
        Write-Host "Database: $database"

        # Update connection string in RDL
        [xml]$rdl = Get-Content $RDLFilePath -ErrorAction Stop
        $DataSources = $rdl.Report.DataSources.DataSource

        if (-not $DataSources) {
            Write-Warning "No DataSources found in RDL. Skipping '$WorkspaceName'."
            $failed++
            continue
        }

        $newConn = "Data Source=$url;Initial Catalog=$database;Authentication=ActiveDirectoryInteractive"
        foreach ($ds in $DataSources) {
            $ds.ConnectionProperties.ConnectString = $newConn
        }

        $rdl.Save($RDLFilePath)
        Write-Host "Connection string updated."

        # Get workspace ID
        $Workspace = Get-PowerBIWorkspace -Name $WorkspaceName -ErrorAction Stop
        if (-not $Workspace) {
            Write-Warning "Workspace '$WorkspaceName' not found. Skipping."
            $failed++
            continue
        }
        Write-Host "Workspace ID: $($Workspace.Id)"

        # Publish
        $id = Publish-RDLFile -RdlFilePath $RDLFilePath -GroupId $Workspace.Id -API $API
        Write-Host "Published '$WorkspaceName' successfully. (Import ID: $id)"
        $success++
    }
    catch {
        Write-Warning "Error processing '$WorkspaceName': $($_.Exception.Message)"
        $failed++
    }
}

Write-Host "`nComplete. Published: $success | Failed: $failed"
Disconnect-PowerBIServiceAccount



