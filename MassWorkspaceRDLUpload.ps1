###UPLOAD ONE POWER BI PAGINATED REPORT TO MULTIPLE WORKSPACES###

###################Variables to to adjust########################
$Tenant = Read-Host "Tenant Name: "
$json = Get-Content "C:\Users\${Tenant}WorkspaceParams.json" | ConvertFrom-Json
$RdlFilePath = "C:\Users\All Users.rdl"

#################################################################
####ASSESS IF POWER BI MODULE IS NEEDED##########################
#################################################################
if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) {
    try {
        Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -ErrorAction Stop
        Write-Host "Power BI Module installed successfully."
    }
    catch {
        Write-Host "Power BI Module installation failed: $($_.Exception.Message)"
    }
}
else {
    Write-Host "Power BI Module already installed. Continuing Script."
}



#################################################################
############CONNECT TO POWER BI WITH SERVICE PRINCIPAL###########
#################################################################

$TenantId = Read-Host "Tenant ID: "
$AppId = Read-Host "App ID: " 
$AppSecret = Read-Host "Secret: "

## Create a secure credential object ## 
$Credential = New-Object PSCredential($AppId, (ConvertTo-SecureString -String $AppSecret -AsPlainText -Force))

## Connect to the Power BI service using the service principal ##
Connect-PowerBIServiceAccount -Tenant $TenantId -ServicePrincipal -Credential $Credential


#########################################################################
########OPTION TO USE SERVICE PRINCIPAL OR OWN POWER BI ACCOUNT##########
#########################################################################

#Connect-PowerBIServiceAccount


#################################################################
###########FUNCTION TO PUBLISH .RDL FILE#########################
#################################################################


##FUNCTION TO UPLOAD REPORTS##
function Publish-ImportRDLFile 
{
    param
    (
        [string]$RdlFilePath,
        [string]$GroupId
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
        $url = "https://api.powerbi.com/v1.0/myorg/groups/$GroupId/imports?datasetDisplayName=$fileName&nameConflict=$nameConflict"
    }
    else {
        $url = "https://api.powerbi.com/v1.0/myorg/imports?datasetDisplayName=$fileName&nameConflict=$nameConflict"
    }

    # Create import
    $report = Invoke-RestMethod -Uri $url -Method Post -Headers $headers -Body $body -ContentType "multipart/form-data"   
    $report.id
}


###############################################################
########REPLACE STRING CONNECTIONS AND PUBLISH#################
###############################################################

#Use $json File to Retrieve Workspace Names, Environment URLs, and Database Names

foreach ($WorkspaceName in $json.PSObject.Properties.Name) {
    $details = $json.$WorkspaceName.updateDetails[0]

    $url = $details.URL
    if (-not $url) { $url = $details.newValue }

    $database = $url -replace "\.crm.*",""

    Write-Host "Name: $WorkspaceName"
    Write-Host "URL: $url"
    Write-Host "Database: $database"

    # Load the RDL as XML
    [xml]$rdl = Get-Content $RdlFilePath
    
    # New connection string
    $newConn = "Data Source=$url;Initial Catalog=$database;Authentication=ActiveDirectoryInteractive"
    
    # Update Data Source (works if only one DataSource)
    $rdl.Report.DataSources.DataSource.ConnectionProperties.ConnectString = $newConn
    
    # Save updated RDL back to file
    $rdl.Save($RdlFilePath)
    Write-Host "Updated $($RdlFilePath)"

    #retrieve workspace
    $GroupId = Get-PowerBIWorkspace -Name $WorkspaceName
    Write-Host "Retrieved GroupID As "$GroupId.Id

    #publish rdl file
    Publish-ImportRDLFile -RdlFilePath $RdlFilePath -GroupId $GroupId.Id
    Write-Host "File has been published to '$WorkspaceName'!"
    Write-Host " "

}



