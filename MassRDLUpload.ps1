#################################################################
# MASS UPLOAD POWER BI PAGINATED (.rdl) REPORTS TO A WORKSPACE
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

$RdlFolder = "C:\Users" #Input RDL Folder Name Here
$Files     = Get-ChildItem -LiteralPath $RdlFolder -Filter *.rdl
$GroupId   = Read-Host "Workspace/Group ID: "
$ServerURL = Read-Host "Server URL: "
$Database = $ServerURL -replace "\.crm.*",""

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
# Connect-PowerBIServiceAccount -Environment $Environment -Tenant $TenantId -ServicePrincipal -Credential $Credential

try {
    Connect-PowerBIServiceAccount -Environment $Environment -ErrorAction Stop
    Write-Host "Connected to Power BI ($ILChoice) successfully."
}
catch {
    Write-Error "Failed to connect to Power BI: $($_.Exception.Message)"
    exit 1
}

#################################################################
# UPDATE CONNECTION STRINGS IN RDL FILES
#################################################################

$NewConn = "Data Source=$ServerURL;Initial Catalog=$Database;Authentication=ActiveDirectoryInteractive"

foreach ($rdlPath in $Files) {
    try {
        [xml]$rdl = Get-Content $rdlPath.FullName -ErrorAction Stop

        $DataSources = $rdl.Report.DataSources.DataSource
        if (-not $DataSources) {
            Write-Warning "No DataSources found in '$($rdlPath.Name)'. Skipping connection update."
            continue
        }

        foreach ($ds in $DataSources) {
            $ds.ConnectionProperties.ConnectString = $NewConn
        }

        $rdl.Save($rdlPath.FullName)
        Write-Host "Connection updated: $($rdlPath.Name)"
    }
    catch {
        Write-Warning "Failed to update connection in '$($rdlPath.Name)': $($_.Exception.Message)"
    }
}

#################################################################
# FUNCTION: PUBLISH RDL FILE TO POWER BI
#################################################################

function Publish-ImportRDLFile 
{
    param
    (
        [string]$RdlFilePath,
        [string]$GroupId,
        [string]$API
    )

    # Get file content and create body
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
# PUBLISH ALL RDL FILES
#################################################################

$success = 0
$failed  = 0

foreach ($file in $Files) {
    try {
        $id = Publish-ImportRDLFile -GroupId $GroupId -RdlFilePath $file.FullName -API $API
        Write-Host "Published: $($file.Name) (Import ID: $id)"
        $success++
    }
    catch {
        Write-Warning "Failed to publish '$($file.Name)': $($_.Exception.Message)"
        $failed++
    }
}

Write-Host "`nComplete. Published: $success | Failed: $failed"
Disconnect-PowerBIServiceAccount