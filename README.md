# PoSHAddPaginatedReportToMultiplePowerBIWorkspaces

PowerShell scripts for bulk-publishing Power BI reports (.rdl paginated reports and .pbix reports) across multiple workspaces, with automatic connection string patching and dataset parameter updates.

---

## Scripts

### [MassRDLUpload.ps1](MassRDLUpload.ps1)

Uploads **all `.rdl` files from a local folder** to a single Power BI workspace, updating each report's connection string before publishing.

**What it does:**
1. Installs the `MicrosoftPowerBIMgmt` module if not already present.
2. Prompts for an Impact Level (Commercial, GCC, GCC High, DOD) and sets the appropriate API endpoint.
3. Prompts for a Workspace/Group ID and a Server URL, then derives the database name from the URL.
4. Connects to Power BI interactively (service principal authentication is included but commented out).
5. Iterates over every `.rdl` file in the configured folder, updates all DataSource connection strings to use `ActiveDirectoryInteractive` auth, and saves the file.
6. Publishes each updated `.rdl` to the target workspace via the Power BI REST API. If a report with the same name already exists it overwrites it; otherwise it aborts on conflict.
7. Prints a final summary of published vs. failed counts and disconnects.

**Key variables to configure before running:**
| Variable | Description |
|---|---|
| `$RdlFolder` | Local folder path containing the `.rdl` files to upload |
| `$GroupId` | Target Power BI Workspace/Group ID (prompted at runtime) |
| `$ServerURL` | Dataverse/SQL Server URL for the connection string (prompted at runtime) |

---

### [MassWorkspaceRDLUpload.ps1](MassWorkspaceRDLUpload.ps1)

Publishes a **single `.rdl` file** to **multiple workspaces**, dynamically patching the connection string for each workspace using a JSON configuration file.

> **Important — JSON file naming:** When prompted for a Tenant Name, you must enter the value that matches the prefix of your workspace params JSON file. The script constructs the path as `{TenantName}WorkspaceParams.json`, so if you enter `Contoso` the script expects a file named `ContosoWorkspaceParams.json` in the script directory. Rename or copy the provided [TenantWorkspaceParams.json](TenantWorkspaceParams.json) template to match before running.

**What it does:**
1. Installs the `MicrosoftPowerBIMgmt` module if not already present.
2. Prompts for an Impact Level and sets the appropriate API endpoint.
3. Prompts for a Tenant Name and constructs the path to `{TenantName}WorkspaceParams.json`.
4. Connects to Power BI interactively (service principal authentication is included but commented out).
5. Loops through each workspace entry in the JSON file:
   - Reads the workspace-specific URL and derives the database name.
   - Updates the `.rdl` file's DataSource connection string to point at that workspace's environment.
   - Looks up the workspace by name to retrieve its ID.
   - Publishes the `.rdl` via the Power BI REST API (overwrites if the report already exists).
6. Prints a final summary of published vs. failed counts and disconnects.

**Key variables to configure before running:**
| Variable | Description |
|---|---|
| `$RDLFilePath` | Full path to the single `.rdl` file to be deployed |
| `{TenantName}WorkspaceParams.json` | Rename this file so its prefix matches what you enter at the Tenant Name prompt |

---

### [MassReportUpload.ps1](MassReportUpload.ps1)

Uploads a **single `.pbix` report** to **multiple workspaces** and updates the dataset's `Environment` parameter in each workspace to point at the correct Dataverse URL.

> **Important — JSON file naming:** Same requirement as above. The script constructs the path as `{TenantName}WorkspaceParams.json` from the Tenant Name prompt, so rename or copy the provided [TenantWorkspaceParams.json](TenantWorkspaceParams.json) template to match your tenant prefix before running.

**What it does:**
1. Prompts for an Impact Level and sets the appropriate API endpoint.
2. Prompts for a Tenant Name and loads the corresponding `{TenantName}WorkspaceParams.json`.
3. Installs the `MicrosoftPowerBIMgmt` module if not already present.
4. Connects to Power BI interactively (service principal authentication is included but commented out).
5. Loops through each workspace entry in the JSON file:
   - Uploads (or overwrites) the `.pbix` file to the workspace using `New-PowerBIReport`.
   - Waits briefly for the dataset to register, then looks it up by name.
   - Calls the Power BI REST API (`Default.UpdateParameters`) to set the `Environment` parameter on the dataset to the workspace-specific URL.
6. Prints completion status per workspace and disconnects.

**Key variables to configure before running:**
| Variable | Description |
|---|---|
| `$ReportName` | File name of the `.pbix` report |
| `$FilePath` | Full local path to the `.pbix` report |
| `{TenantName}WorkspaceParams.json` | Rename this file so its prefix matches what you enter at the Tenant Name prompt |

---

## Workspace Parameters JSON

[MassWorkspaceRDLUpload.ps1](MassWorkspaceRDLUpload.ps1) and [MassReportUpload.ps1](MassReportUpload.ps1) both rely on a JSON file named `{TenantName}WorkspaceParams.json`. The file maps Power BI workspace display names to their Dataverse environment URLs.

**Naming convention:** The filename prefix must exactly match the value you enter at the `Tenant Name` prompt. For example, entering `Contoso` at the prompt means the script will look for `ContosoWorkspaceParams.json`.

**Template ([TenantWorkspaceParams.json](TenantWorkspaceParams.json)):**
```json
{
  "PowerBI Workspace Name 1": {
    "updateDetails": [
      { "URL": "org1.crm.dynamics.com" }
    ]
  },
  "PowerBI Workspace Name 2": {
    "updateDetails": [
      { "URL": "org2.crm.dynamics.com" }
    ]
  },
  "Power BI Workspace Name 3": {
    "updateDetails": [
      { "URL": "org3.crm.dynamics.com" }
    ]
  }
}
```

- **Key** — must match the exact display name of the Power BI workspace.
- **`updateDetails[0].URL`** — the Dataverse environment URL used to build the connection string and/or dataset parameter.

Add or remove workspace entries as needed. Each workspace that should receive the report needs its own key/URL pair.

---

## Prerequisites

- PowerShell 5.1 or later
- `MicrosoftPowerBIMgmt` module (auto-installed by the scripts if missing)
- An account with **Contributor** or higher access to the target Power BI workspaces
- Network access to the appropriate Power BI API endpoint for your environment

## Supported Environments

| Selection | API Endpoint | PowerBI Environment |
|---|---|---|
| Commercial | `api.powerbi.com` | `Public` |
| GCC | `api.powerbigov.us` | `USGov` |
| GCC High | `api.high.powerbigov.us` | `USGovHigh` |
| DOD | `api.mil.powerbigov.us` | `USGovMil` |

## Optional: Service Principal Authentication

All three scripts include commented-out blocks for service principal (non-interactive) authentication. To use a service principal, uncomment and populate the following in whichever script you are running:

```powershell
$TenantId  = Read-Host "Tenant ID"
$AppId     = Read-Host "App ID"
$AppSecret = Read-Host "Secret"
$Credential = New-Object PSCredential($AppId, (ConvertTo-SecureString -String $AppSecret -AsPlainText -Force))
Connect-PowerBIServiceAccount -Environment $Environment -Tenant $TenantId -ServicePrincipal -Credential $Credential
```
