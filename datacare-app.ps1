<#
.SYNOPSIS
    Microsoft 365 Usage Data Collector and SQL Server Loader

.DESCRIPTION
    This script authenticates to Microsoft 365 and Exchange Online using application permissions
    and retrieves usage and activity details.

    Retrieved data includes:
        - Exchange mailbox usage
        https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0&tabs=http
        https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/get-exomailboxstatistics?view=exchange-ps
        - OneDrive usage
        https://learn.microsoft.com/en-us/graph/api/reportroot-getonedriveusageaccountdetail?view=graph-rest-1.0&tabs=http
        - SharePoint site usage
        https://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail?view=graph-rest-1.0&tabs=http
        - Azure AD users
        https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http

    All data is:
        - Logged with structured execution tracking
        - Enriched with metadata (period, execution date, source report)
        - Stored into SQL Server tables (auto-created if not existing)
        - Tracked in dbo.ExecutionLog for auditing

    The script implements:
        - OAuth2 client credentials flow
        - Structured logging with timestamp
        - Automatic modules installations

.REQUIREMENTS
    - PowerShell 5.1 or higher
    - Internet access to:
        * https://login.microsoftonline.com
        * https://graph.microsoft.com
        * reports*.office.com (Graph reports redirect)
    - SQL Server: server localhost\SQLEXPRESS with: enctyption and TCP/IP network protocol "Enabled"
    - Azure AD App Registration (Application permissions) with:
        * Reports.Read.All (Microsoft Graph) - grant admin consent
        * User.Read.All (Microsoft Graph) - grant admin consent
        * Exchange.ManageAsAp (Office 365Exchange Online(1)) - grant admin consent

.CONFIGURATION
    The following variables must be configured:
    - Azure AD:
        $TenantId
        $ClientId
        $ClientSecret
    - Execution:
        $Department

.EXECUTION
    1. Open PowerShell with admin privileges
    2. Navigate to the script folder
    3. Set SecretManagement and SecretStore properties:
        3.1 Register-SecretVault `
                -Name LocalVault `
                -ModuleName Microsoft.PowerShell.SecretStore `
                -DefaultVault
        3.2 Set-Secret -Name GraphClientSecret -Secret "my-secret"   
    4. Run the command: .\datacare-app.ps1
    5. Enter the required credentials

.COPYRIGHT
    © 2026 Business Integration Partners. All rights reserved.

.LICENSE
    Internal corporate use only. Unauthorized distribution or modification is prohibited.

.VERSION
    1.0.0
#>

# ======================
#     CONFIGURATIONS
# ======================
$Config = @{
    TenantId     = "76ff1baa-3307-46aa-a752-cc3736d8a2b2" #your_tenantId
    ClientId     = "dd80738f-6094-43ff-bf26-03fe4e3bc7da" #your_clientId

    Sql = @{
        Server      = "localhost\SQLEXPRESS"
        SqlDBMaster = "master"
        SqlDBTarget = "DataCare"
        CreateTable_Exchange = "
        IF OBJECT_ID('dbo.Exchange','U') IS NULL
        CREATE TABLE dbo.Exchange (
            StorageUsedGB FLOAT,
            ___Report_Refresh_Date NVARCHAR(50),
            User_Principal_Name NVARCHAR(255) NOT NULL,
            Display_Name NVARCHAR(255),
            Is_Deleted NVARCHAR(50),
            Deleted_Date NVARCHAR(50),
            Created_Date NVARCHAR(50),
            Last_Activity_Date NVARCHAR(50),
            Item_Count INT,
            Storage_Used__Byte_ BIGINT,
            Issue_Warning_Quota__Byte_ BIGINT,
            Prohibit_Send_Quota__Byte_ BIGINT,
            Prohibit_Send_Receive_Quota__Byte_ BIGINT,
            Deleted_Item_Count INT,
            Deleted_Item_Size__Byte_ BIGINT,
            Deleted_Item_Quota__Byte_ BIGINT,
            Has_Archive NVARCHAR(50),
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100),
            Department NVARCHAR(50),

            -- Primary mailbox statistics
            Primary_Item_Count INT,
            Primary_TotalItemSize NVARCHAR(50),
            Primary_Total_Size_Bytes BIGINT,
            Primary_SystemMessage_Count INT,
            Primary_SystemMessage_Size_Bytes BIGINT,
            Primary_Recoverable_Count INT,
            Primary_Recoverable_Size_Bytes BIGINT,
            Primary_Recoverable_Mode NVARCHAR(50),

            -- Archive mailbox statistics
            Archive_Item_Count INT,
            Archive_TotalItemSize NVARCHAR(50),
            Archive_Total_Size_Bytes BIGINT,
            Archive_SystemMessage_Count INT,
            Archive_SystemMessage_Size_Bytes BIGINT,
            Archive_Recoverable_Count INT,
            Archive_Recoverable_Size_Bytes BIGINT,
            Archive_Recoverable_Mode NVARCHAR(50)
        );"
        CreateTable_OneDrive = "
        IF OBJECT_ID('dbo.OneDrive','U') IS NULL
        CREATE TABLE dbo.OneDrive (
            StorageUsedGB FLOAT,
            ___Report_Refresh_Date NVARCHAR(50),
            Site_Id NVARCHAR(255),
            Site_URL NVARCHAR(500),
            Owner_Display_Name NVARCHAR(255),
            Is_Deleted NVARCHAR(50),
            Last_Activity_Date NVARCHAR(50),
            File_Count INT,
            Active_File_Count INT,
            Storage_Used__Byte_ BIGINT,
            Storage_Allocated__Byte_ BIGINT,
            Owner_Principal_Name NVARCHAR(255) NOT NULL,
            Department NVARCHAR(50),
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100)
        );"
        CreateTable_SharePoint = "
        IF OBJECT_ID('dbo.SharePoint','U') IS NULL
        CREATE TABLE dbo.SharePoint (
            StorageUsedGB FLOAT,
            ___Report_Refresh_Date NVARCHAR(50),
            Site_Id NVARCHAR(255),
            Site_URL NVARCHAR(500),
            Owner_Display_Name NVARCHAR(255),
            Is_Deleted NVARCHAR(50),
            Last_Activity_Date NVARCHAR(50),
            File_Count INT,
            Active_File_Count INT,
            Page_View_Count INT,
            Visited_Page_Count INT,
            Storage_Used__Byte_ BIGINT,
            Storage_Allocated__Byte_ BIGINT,
            Root_Web_Template NVARCHAR(100),
            Owner_Principal_Name NVARCHAR(255) NOT NULL,
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100)
        );"
        CreateTable_Users = "
        IF OBJECT_ID('dbo.Users','U') IS NULL
        CREATE TABLE dbo.Users (
            Id NVARCHAR(255),
            DisplayName NVARCHAR(255),
            UserPrincipalName NVARCHAR(255) NOT NULL,
            Mail NVARCHAR(255),
            Department NVARCHAR(255),
            JobTitle NVARCHAR(255),
            AccountEnabled NVARCHAR(50),
            CreatedDateTime NVARCHAR(50),
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100)
        );"   
        CreateTable_ExecutionLog = "
        IF OBJECT_ID('dbo.ExecutionLog','U') IS NULL
        CREATE TABLE dbo.ExecutionLog (
            ExecutionId UNIQUEIDENTIFIER,
            ExecutionDate DATETIME2,
            ReportName NVARCHAR(100),
            Status NVARCHAR(50),
            RowsRetrieved INT,
            RowsInserted INT,
            DurationSeconds INT,
            ErrorMessage NVARCHAR(MAX),
            MachineName NVARCHAR(255),
            PowerShellVersion NVARCHAR(50)
        );"   
        CreateTable_PowerBIDataModel = "
        IF OBJECT_ID('dbo.PowerBIDataModel','U') IS NULL
        CREATE TABLE dbo.PowerBIDataModel (
            Exchange_Total_Primary_Item_Count INT,
            Exchange_Total_Archive_Item_Count INT,
            Exchange_Total_Primary_Total_Size_GB DECIMAL(18,2),
            Exchange_Total_Archive_Total_Size_GB DECIMAL(18,2),
            OneDrive_Total_File_Count INT,
            OneDrive_Total_StorageUsedGB FLOAT,
            SharePoint_Total_File_Count INT,
            SharePoint_Total_StorageUsedGB FLOAT,
            Users_Total INT
        );"
    }

    Execution = @{
        Period     = "D180"
        Department = "Information Technology"
    }
}
$masterConnectionString = "Server=$($Config.Sql.Server);Database=$($Config.Sql.SqlDBMaster);Trusted_Connection=True;TrustServerCertificate=True;"
$targetConnectionString = "Server=$($Config.Sql.Server);Database=$($Config.Sql.SqlDBTarget);Trusted_Connection=True;TrustServerCertificate=True;"


# ======================
#      FUNCTIONS
# ======================
function Write-Log {
     param (
        [Parameter(Mandatory)] [string]$Message,
        [ConsoleColor]$ForegroundColor = [ConsoleColor]::White
    )

    $dateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$dateTime - $Message"
    [System.IO.File]::AppendAllText("$PSScriptRoot\DataCare.log", "$line`r`n")
    Write-Host $line -ForegroundColor $ForegroundColor
}

function Write-ExecutionLog {
    param(
        [guid]$ExecutionId,
        [string]$ReportName,
        [string]$Status,
        [int]$RowsRetrieved = 0,
        [int]$RowsInserted = 0,
        [int]$DurationSeconds = 0,
        [string]$ErrorMessage = $null
    )

    $MachineName = $env:COMPUTERNAME
    $PSVersion   = $PSVersionTable.PSVersion.ToString()

    if ($ReportName) { $ReportName = $ReportName.Replace("'", "''") } else { $ReportName = "" }
    if ($Status)     { $Status     = $Status.Replace("'", "''") } else { $Status = "" }
    if ($ErrorMessage) { $ErrorMessage = $ErrorMessage.Replace("'", "''") } else { $ErrorMessage = $null }

    $query = @"
INSERT INTO dbo.ExecutionLog
(ExecutionId, ExecutionDate, ReportName, Status,
 RowsRetrieved, RowsInserted, DurationSeconds,
 ErrorMessage, MachineName, PowerShellVersion)
VALUES
('$ExecutionId', SYSDATETIME(), '$ReportName', '$Status',
 $RowsRetrieved, $RowsInserted, $DurationSeconds,
 $(if($ErrorMessage){"'$ErrorMessage'"}else{"NULL"}),
 '$MachineName', '$PSVersion')
"@

    try {
        Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $query
    }
    catch {
        Write-Host "Failed to write ExecutionLog: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Import-RequiredModule {
    param (
        [Parameter(Mandatory)]
        [string]$ModuleName,

        [ValidateSet("CurrentUser","AllUsers")]
        [string]$Scope = "CurrentUser"
    )

    try {
        if (Get-Module -Name $ModuleName) {
            Write-Log "Module $ModuleName already loaded" -ForegroundColor Green
            return
        }

        if (Get-Module -ListAvailable -Name $ModuleName) {
            Import-Module $ModuleName -Force -ErrorAction Stop
            Write-Log "Module $ModuleName imported" -ForegroundColor Green
            return
        }

        Write-Log "Module $ModuleName not found. Installing..." -ForegroundColor Yellow

        Install-Module `
            -Name $ModuleName `
            -Scope $Scope `
            -AllowClobber `
            -Force `
            -ErrorAction Stop

        Import-Module $ModuleName -Force -ErrorAction Stop

        Write-Log "Module $ModuleName installed and imported successfully" -ForegroundColor Green
    }
    catch {
        Write-Log "Failed to load module $ModuleName : $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# EXCHANGE
function Connect-ExchangeAppOnly {
    Write-Log "Connecting to Exchange Online ..." -ForegroundColor Cyan
    Write-Log "Enter/Select your credentials to log in (MFA)..." -ForegroundColor Magenta

    Connect-ExchangeOnline

    Write-Log "Connected to Exchange Online" -ForegroundColor Green
}

function Convert-ToBytes {
    param([string]$SizeString)

    if (-not $SizeString) { return 0 }
    if ($SizeString -match '\((\d[\d,]*) bytes\)') {
        return [Int64]($matches[1] -replace ',', '')
    }
    elseif ($SizeString -match '([0-9,.]+)\s*GB') {
        return [math]::Round([double]$matches[1] * 1GB)
    }
    elseif ($SizeString -match '([0-9,.]+)\s*MB') {
        return [math]::Round([double]$matches[1] * 1MB)
    }
    elseif ($SizeString -match '([0-9,.]+)\s*KB') {
        return [math]::Round([double]$matches[1] * 1KB)
    }
    else {
        return 0
    }
}

function Get-ExchangeMailboxDeepStats {
    param (
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $result = [ordered]@{
        Primary_Item_Count               = 0
        Primary_Total_Size_Bytes         = 0
        Primary_TotalItemSize            = ""
        Primary_SystemMessageCount       = 0
        Primary_SystemMessageSize        = ""
        Primary_Recoverable_Count        = 0
        Primary_Recoverable_Size_Bytes   = 0
        Primary_Recoverable_Mode         = "NotPresent"

        Archive_Item_Count               = 0
        Archive_Total_Size_Bytes         = 0
        Archive_TotalItemSize            = ""
        Archive_SystemMessageCount       = 0
        Archive_SystemMessageSize        = ""
        Archive_Recoverable_Count        = 0
        Archive_Recoverable_Size_Bytes   = 0
        Archive_Recoverable_Mode         = "NotPresent"
    }

    try {
        $primaryStats = Get-EXOMailboxStatistics `
            -Identity $UserPrincipalName `
            -Properties ItemCount,TotalItemSize,SystemMessageCount,SystemMessageSize `
            -ErrorAction Stop

        if ($primaryStats) {

            $result.Primary_Item_Count         = $primaryStats.ItemCount
            $result.Primary_TotalItemSize      = $primaryStats.TotalItemSize
            $result.Primary_SystemMessageCount = $primaryStats.SystemMessageCount
            $result.Primary_SystemMessageSize  = $primaryStats.SystemMessageSize

            if ($primaryStats.TotalItemSize) {
                $result.Primary_Total_Size_Bytes =
                    Convert-ToBytes $primaryStats.TotalItemSize
            }
        }

        $primaryRI = Get-MailboxFolderStatistics `
            -Identity $UserPrincipalName `
            -FolderScope RecoverableItems `
            -ErrorAction SilentlyContinue

        if ($primaryRI) {
            foreach ($folder in $primaryRI) {
                if ($folder.Name -eq "Recoverable Items") {

                    $result.Primary_Recoverable_Count =
                        $folder.ItemsInFolderAndSubfolders

                    if ($folder.FolderAndSubfolderSize) {
                        $result.Primary_Recoverable_Size_Bytes =
                            Convert-ToBytes $folder.FolderAndSubfolderSize
                    }

                    $result.Primary_Recoverable_Mode = "Aggregated"
                    break
                }
            }
        }

        $archiveStats = Get-EXOMailboxStatistics `
            -Identity $UserPrincipalName `
            -Archive `
            -Properties ItemCount,TotalItemSize,SystemMessageCount,SystemMessageSize `
            -ErrorAction SilentlyContinue

        if ($archiveStats) {

            $result.Archive_Item_Count         = $archiveStats.ItemCount
            $result.Archive_TotalItemSize      = $archiveStats.TotalItemSize
            $result.Archive_SystemMessageCount = $archiveStats.SystemMessageCount
            $result.Archive_SystemMessageSize  = $archiveStats.SystemMessageSize

            if ($archiveStats.TotalItemSize) {
                $result.Archive_Total_Size_Bytes =
                    Convert-ToBytes $archiveStats.TotalItemSize
            }

            $archiveRI = Get-MailboxFolderStatistics `
                -Identity $UserPrincipalName `
                -Archive `
                -FolderScope RecoverableItems `
                -ErrorAction SilentlyContinue

            if ($archiveRI) {
                foreach ($folder in $archiveRI) {
                    if ($folder.Name -eq "Recoverable Items") {

                        $result.Archive_Recoverable_Count =
                            $folder.ItemsInFolderAndSubfolders

                        if ($folder.FolderAndSubfolderSize) {
                            $result.Archive_Recoverable_Size_Bytes =
                                Convert-ToBytes $folder.FolderAndSubfolderSize
                        }

                        $result.Archive_Recoverable_Mode = "Aggregated"
                        break
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Mailbox enrichment failed for $UserPrincipalName : $($_.Exception.Message)" -ForegroundColor Red
    }

    return $result
}

# SQLSERVER
function Test-SqlConnection {
    try {
        Invoke-Sqlcmd -ServerInstance $Config.Sql.Server -Database $Config.Sql.SqlDBTarget
        Write-Log "SQL Server connection successful" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Log "Cannot connect to SQL Server: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Initialize-Database {
    Write-Log "Ensuring database $($Config.Sql.SqlDBTarget) exists..." Cyan

    $createDbQuery = @"
IF DB_ID(N'$($Config.Sql.SqlDBTarget)') IS NULL
    CREATE DATABASE [$($Config.Sql.SqlDBTarget)];
"@

    Invoke-Sqlcmd -ConnectionString $masterConnectionString -Query $createDbQuery

    Write-Log "Database verified/created" Green

    $tables = @{
        ExecutionLog = $Config.Sql.CreateTable_ExecutionLog
        Users        = $Config.Sql.CreateTable_Users
        Exchange     = $Config.Sql.CreateTable_Exchange
        OneDrive     = $Config.Sql.CreateTable_OneDrive
        SharePoint   = $Config.Sql.CreateTable_SharePoint
    }

    foreach ($table in $tables.GetEnumerator()) {
        $tableName = $table.Key
        $createSql = $table.Value

        $checkQuery = "IF OBJECT_ID('dbo.$tableName','U') IS NULL SELECT 0 ELSE SELECT 1"

        $exists = Invoke-Sqlcmd `
            -ConnectionString $targetConnectionString `
            -Query $checkQuery |
            Select-Object -ExpandProperty Column1

        if ($exists -eq 1) {
            Write-Log "Table '$tableName' already exists"
        }
        else {
            Invoke-Sqlcmd `
                -ConnectionString $targetConnectionString `
                -Query $createSql

            Write-Log "Table '$tableName' created successfully" Green
        }
    }
    Write-Log "Database initialization completed" Green
}

function NormalizeData {
    param([string]$Attribute)
    if ([string]::IsNullOrEmpty($Attribute)) { return "" }
    return ($Attribute.ToLower() -replace '[^a-z0-9]')
}

function Write-ToSqlTable {
    param (
        [Parameter(Mandatory)][string]$TableName,
        [Parameter(Mandatory)][array]$Data
    )

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Log "No data to write for table '$TableName'" Yellow
        return
    }

    try {
        $connection = New-Object System.Data.SqlClient.SqlConnection($targetConnectionString)
        $connection.Open()

        $schemaQuery = "
        SELECT COLUMN_NAME 
        FROM INFORMATION_SCHEMA.COLUMNS 
        WHERE TABLE_NAME = '$TableName'
        "

        $sqlColumns = Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $schemaQuery |
                      Select-Object -ExpandProperty COLUMN_NAME

        if (-not $sqlColumns) {
            throw "Table dbo.$TableName does not exist."
        }

        $table = New-Object System.Data.DataTable
        foreach ($col in $sqlColumns) {
            $null = $table.Columns.Add($col)
        }

        $propertyMap = @{}
        foreach ($prop in $Data[0].PSObject.Properties) {
            $normalized = NormalizeData $prop.Name
            if (-not $propertyMap.ContainsKey($normalized)) {
                $propertyMap[$normalized] = $prop.Name
            }
        }

        foreach ($row in $Data) {
            $dr = $table.NewRow()

            foreach ($sqlCol in $sqlColumns) {
                $normalizedSql = $sqlCol.Trim().ToLower()

                if ($normalizedSql -eq "insertedat") {
                    $dr[$sqlCol] = Get-Date
                    continue
                }
                if ($normalizedSql -eq "sourcereport") {
                    $dr[$sqlCol] = $TableName
                    continue
                }
                if ($normalizedSql -eq "reportdate") {
                    $dr[$sqlCol] = Get-Date
                    continue
                }
                if ($normalizedSql -eq "storageusedgb") {

                    if ($propertyMap.ContainsKey("storageusedbyte")) {
                        $bytes = $row.($propertyMap["storageusedbyte"])
                        $dr[$sqlCol] = if ($bytes) { [math]::Round(($bytes / 1GB),2) } else { [DBNull]::Value }
                    }
                    else {
                        $dr[$sqlCol] = [DBNull]::Value
                    }
                    continue
                }
                if ($TableName -eq "SharePoint" -and $normalizedSql -eq "owner_principal_name") {
                    if ($propertyMap.ContainsKey("ownerprincipalname")) {
                        $value = $row.($propertyMap["ownerprincipalname"])
                        $dr[$sqlCol] = if ($value) { $value } else { "Value Not Present" }
                    }
                    else {
                        $dr[$sqlCol] = "Value Not Present"
                    }
                    continue
                }

                $normalizedTarget = NormalizeData $sqlCol
                if ($propertyMap.ContainsKey($normalizedTarget)) {
                    $value = $row.($propertyMap[$normalizedTarget])
                    if ($null -eq $value -or $value -eq "") {
                        $dr[$sqlCol] = [DBNull]::Value
                    }
                    else {
                        $dr[$sqlCol] = $value
                    }
                }
                else {
                    $dr[$sqlCol] = [DBNull]::Value
                }
            }
            $table.Rows.Add($dr)
        }

        $bulkCopy = New-Object System.Data.SqlClient.SqlBulkCopy($connection)
        $bulkCopy.DestinationTableName = "dbo.$TableName"
        $bulkCopy.BatchSize = 5000
        $bulkCopy.BulkCopyTimeout = 0

        foreach ($col in $sqlColumns) {
            $null = $bulkCopy.ColumnMappings.Add($col,$col)
        }

        $bulkCopy.WriteToServer($table)
        $connection.Close()

        Write-Log "Inserted $($table.Rows.Count) rows into $TableName successfully" Green
    }
    catch {
        Write-Log "ERROR writing to $TableName : $($_.Exception.Message)" Red
        throw
    }
}

function Get-ReportCountFromDb {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ReportName
    )

    try {
        $tableMap = @{
            "Exchange"    = "dbo.Exchange"
            "OneDrive"    = "dbo.OneDrive"
            "SharePoint"  = "dbo.SharePoint"
            "Users"       = "dbo.Users"
        }

        if (-not $tableMap.ContainsKey($ReportName)) {
            throw "Unknown report name: $ReportName"
        }

        $tableName = $tableMap[$ReportName]
        $query = "SELECT COUNT(*) AS Total FROM $tableName"
        $result = Invoke-Sqlcmd `
            -ConnectionString $targetConnectionString `
            -Query $query
        return [int]$result.Total
    }
    catch {
        Write-Log "Failed to retrieve $ReportName count from DB: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# ENTRAID
function Get-GraphAccessToken {
    Write-Log "Requesting Microsoft Graph token..." Cyan

    $body = @{
        client_id     = $Config.ClientId
        client_secret = $Config.ClientSecret
        scope         = "https://graph.microsoft.com/.default"
        grant_type    = "client_credentials"
    }

    $uri = "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token"
    try {
        $response = Invoke-RestMethod `
            -Uri $uri `
            -Method POST `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded"

        Write-Log "Graph token acquired" Green
        return $response.access_token
    }
    catch {
        throw "Graph authentication failed: $($_.Exception.Message)"
    }
}

function Invoke-GraphRequest {
    param (
        [string]$Url,
        [hashtable]$Headers
    )

    Write-Log "Calling: $Url" -ForegroundColor Yellow

    try { return Invoke-RestMethod -Uri $Url -Headers $Headers -Method GET }
    catch {
        Write-Log "Request failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}


# ======================
#           MAIN
# ======================
try {
    Write-Log "=== START DATACARE EXE ==="
    
    $ExecutionId = [guid]::NewGuid()

    Import-RequiredModule -ModuleName "SqlServer"
    Import-RequiredModule -ModuleName "ExchangeOnlineManagement"
    Import-RequiredModule -ModuleName "Microsoft.PowerShell.SecretManagement"
    Import-RequiredModule -ModuleName "Microsoft.PowerShell.SecretStore"

    if (-not (Test-SqlConnection)) {
        Write-Log "SQL connection failed." -ForegroundColor Red
        throw "SQL connection failed."
    }

    Initialize-Database

    $TotalRowsRetrieved = 0
    $TotalRowsInserted  = 0

    $ReportTables = @{
        Exchange   = $Config.Sql.CreateTable_Exchange
        OneDrive   = $Config.Sql.CreateTable_OneDrive
        SharePoint = $Config.Sql.CreateTable_SharePoint
        Users      = $Config.Sql.CreateTable_Users
        DataModel  = $Config.Sql.CreateTable_PowerBIDataModel
    }
    foreach ($report in $ReportTables.GetEnumerator()) {
        $reportName = $report.Key
        $createQuery = $report.Value
        $tableName = "dbo.$reportName"

        $count = Get-ReportCountFromDb -ReportName $reportName
        if ($count -gt 0) {
            Write-Log "Number of records in $tableName : $count. Dropping and recreating $tableName table..." Yellow
            $dropQuery = "IF OBJECT_ID('$tableName','U') IS NOT NULL DROP TABLE $tableName;"
            Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $dropQuery
            Write-Log "$tableName table dropped successfully." Green

            Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $createQuery
            Write-Log "Table '$tableName' created successfully" Green
        }
    }

    $Config.ClientSecret = Get-Secret GraphClientSecret -AsPlainText
    $AccessToken = Get-GraphAccessToken
    $ReportHeaders = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "text/csv"
    }
    $UserHeaders = @{
        Authorization    = "Bearer $AccessToken"
        ConsistencyLevel = "eventual"
    }

    Connect-ExchangeAppOnly

    $TaskStart = Get-Date

    #STEP 1: exchange, onedrive and sharepoint
    $Reports = @{
        Exchange   = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$($Config.Execution.Period)')"
        OneDrive   = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$($Config.Execution.Period)')"
        SharePoint = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='$($Config.Execution.Period)')"
    }

    foreach ($ReportName in $Reports.Keys) {
        Write-Log "STEP 1 - CASE: $ReportName" -ForegroundColor Cyan
        try {
            $Response = Invoke-GraphRequest -Url $Reports[$ReportName] -Headers $ReportHeaders
            if ($Response) {
                $Data = $Response | ConvertFrom-Csv
                if ($Data -and $Data.Count -gt 0) {
                    $RowsRetrievedExchange = 0
                    $RowsRetrievedOneDrive = 0
                    $RowsInserted  = 0
                    if ($ReportName -eq "Exchange") {
                        foreach ($Row in $Data) {
                            $ReportRefreshDate = ($Row.PSObject.Properties |
                                Where-Object { $_.Name -like "*Report Refresh Date*" }).Name
                            $RefreshDate = $Row.$ReportRefreshDate

                            $UserPrincipalName = $Row.'User Principal Name'.Trim() -replace "^[\uFEFF]", ""
                            if ([string]::IsNullOrEmpty($UserPrincipalName)) {
                                Write-Log "UPN is empty" -ForegroundColor Yellow
                                continue
                            }
                            $EncodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)
                            $Url = "https://graph.microsoft.com/v1.0/users/"+$EncodedUpn+"?`$select=department"

                            try {
                                $Response = Invoke-GraphRequest -Url $Url -Headers $UserHeaders
                                $UserDepartment = $Response.department

                                if ($UserDepartment -eq $Config.Execution.Department) {
                                    Write-Host "User $UserPrincipalName in department: $($Config.Execution.Department)"

                                    try {
                                        $deepStats = Get-ExchangeMailboxDeepStats -UserPrincipalName $UserPrincipalName

                                        $exchangeObject = [PSCustomObject]@{
                                            displayName                 = $Row.'Display Name'
                                            userPrincipalName           = $Row.'User Principal Name'
                                            mail                        = $Row.'User Principal Name'
                                            department                  = $Config.Execution.Department
                                            Report_Refresh_Date         = $RefreshDate
                                            Is_Deleted                  = $Row.'Is Deleted'
                                            Deleted_Date                = $Row.'Deleted Date'
                                            Created_Date                = $Row.'Created Date'
                                            Last_Activity_Date          = $Row.'Last Activity Date'
                                            Item_Count                  = $Row.'Item Count'
                                            Storage_Used__Byte_         = $Row.'Storage Used (Byte)'
                                            Issue_Warning_Quota__Byte_  = $Row.'Issue Warning Quota (Byte)'
                                            Prohibit_Send_Quota__Byte_  = $Row.'Prohibit Send Quota (Byte)'
                                            Prohibit_Send_Receive_Quota__Byte_ = $Row.'Prohibit Send/Receive Quota (Byte)'
                                            Deleted_Item_Count          = $Row.'Deleted Item Count'
                                            Deleted_Item_Size__Byte_    = $Row.'Deleted Item Size (Byte)'
                                            Deleted_Item_Quota__Byte_   = $Row.'Deleted Item Quota (Byte)'
                                            Has_Archive                 = $Row.'Has Archive'
                                            Report_Period               = $Row.'Report Period'
                                        }
                                        foreach ($key in $deepStats.Keys) {
                                            $exchangeObject | Add-Member -NotePropertyName $key -NotePropertyValue $deepStats[$key] -Force
                                        }

                                        Write-Log "Writing $ReportName record into SQLServer ..." -ForegroundColor Cyan
                                        Write-ToSqlTable -TableName "Exchange" -Data @($exchangeObject)
                                        $RowsInserted++
                                    }
                                    catch {
                                        Write-Log "Deep stats failed for $UserPrincipalName : $($_.Exception.Message)" -ForegroundColor Yellow
                                    }
                                }
                                else {
                                    Write-Host "Users $UserPrincipalName not in department: $($Config.Execution.Department)"
                                    $OtherDepartment = "Other"

                                    $exchangeObject = [PSCustomObject]@{
                                            displayName                 = $Row.'Display Name'
                                            userPrincipalName           = $Row.'User Principal Name'
                                            mail                        = $Row.'User Principal Name'
                                            department                  = $OtherDepartment
                                            Report_Refresh_Date         = $RefreshDate
                                            Is_Deleted                  = $Row.'Is Deleted'
                                            Deleted_Date                = $Row.'Deleted Date'
                                            Created_Date                = $Row.'Created Date'
                                            Last_Activity_Date          = $Row.'Last Activity Date'
                                            Item_Count                  = $Row.'Item Count'
                                            Storage_Used__Byte_         = $Row.'Storage Used (Byte)'
                                            Issue_Warning_Quota__Byte_  = $Row.'Issue Warning Quota (Byte)'
                                            Prohibit_Send_Quota__Byte_  = $Row.'Prohibit Send Quota (Byte)'
                                            Prohibit_Send_Receive_Quota__Byte_ = $Row.'Prohibit Send/Receive Quota (Byte)'
                                            Deleted_Item_Count          = $Row.'Deleted Item Count'
                                            Deleted_Item_Size__Byte_    = $Row.'Deleted Item Size (Byte)'
                                            Deleted_Item_Quota__Byte_   = $Row.'Deleted Item Quota (Byte)'
                                            Has_Archive                 = $Row.'Has Archive'
                                            Report_Period               = $Row.'Report Period'
                                        }

                                    Write-Log "Writing $ReportName record into SQLServer ..." -ForegroundColor Cyan
                                    Write-ToSqlTable -TableName $ReportName -Data @($exchangeObject)   
                                    $RowsInserted++
                                }
                            } 
                            catch {
                                Write-Log "Error while recovering department for $UserPrincipalName : $_" -ForegroundColor Yellow
                            }
                        }
                        $RowsRetrievedExchange = $Data.Count

                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status "SUCCESS" `
                            -RowsRetrieved $RowsRetrievedExchange `
                            -RowsInserted $RowsInserted `
                            -DurationSeconds ([int]((Get-Date) - $TaskStart).TotalSeconds)
                    }
                    elseif ($ReportName -eq "OneDrive") {
                        foreach ($Row in $Data) {
                            $ReportRefreshDate = ($Row.PSObject.Properties |
                                Where-Object { $_.Name -like "*Report Refresh Date*" }).Name
                            $RefreshDate = $Row.$ReportRefreshDate

                            $UserPrincipalName = $Row.'Owner Principal Name'.Trim() -replace "^[\uFEFF]", ""
                            if ([string]::IsNullOrEmpty($UserPrincipalName)) {
                                Write-Log "UPN is empty" -ForegroundColor Yellow
                                continue
                            }
                            $EncodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)
                            $Url = "https://graph.microsoft.com/v1.0/users/"+$EncodedUpn+"?`$select=department"

                            try {
                                $Response = Invoke-GraphRequest -Url $Url -Headers $UserHeaders
                                $UserDepartment = $Response.department

                                if ($UserDepartment -eq $Department) {
                                    Write-Host "Users $UserPrincipalName not in department: $($Config.Execution.Department)"
                                    $OtherDepartment = "Other"

                                    $oneDriveObject = [PSCustomObject]@{
                                        StorageUsedGB            = $StorageUsedGB
                                        ___Report_Refresh_Date   = $RefreshDate
                                        Site_Id                  = $Row.'Site Id'
                                        Site_URL                 = $Row.'Site URL'
                                        Owner_Display_Name       = $Row.'Owner Display Name'
                                        Is_Deleted               = $Row.'Is Deleted'
                                        Last_Activity_Date       = $Row.'Last Activity Date'
                                        File_Count               = $Row.'File Count'
                                        Active_File_Count        = $Row.'Active File Count'
                                        Storage_Used__Byte_      = $Row.'Storage Used (Byte)'
                                        Storage_Allocated__Byte_ = $Row.'Storage Allocated (Byte)'
                                        Owner_Principal_Name     = $Row.'Owner Principal Name'
                                        Report_Period            = $Row.'Report Period'
                                        department               = $Config.Execution.Department
                                        ReportPeriod             = $Period
                                        ReportDate               = $RefreshDate
                                        InsertedAt               = (Get-Date)
                                        SourceReport             = $ReportName    
                                        }

                                    Write-Log "Writing $ReportName record into SQLServer ..." -ForegroundColor Cyan
                                    Write-ToSqlTable -TableName $ReportName -Data @($oneDriveObject) 
                                    $RowsInserted++
                                }
                                else {
                                    Write-Host "Users $UserPrincipalName not in department: $($Config.Execution.Department)"
                                    $OtherDepartment = "Other"

                                    $oneDriveObject = [PSCustomObject]@{
                                        StorageUsedGB            = $StorageUsedGB
                                        ___Report_Refresh_Date   = $RefreshDate
                                        Site_Id                  = $Row.'Site Id'
                                        Site_URL                 = $Row.'Site URL'
                                        Owner_Display_Name       = $Row.'Owner Display Name'
                                        Is_Deleted               = $Row.'Is Deleted'
                                        Last_Activity_Date       = $Row.'Last Activity Date'
                                        File_Count               = $Row.'File Count'
                                        Active_File_Count        = $Row.'Active File Count'
                                        Storage_Used__Byte_      = $Row.'Storage Used (Byte)'
                                        Storage_Allocated__Byte_ = $Row.'Storage Allocated (Byte)'
                                        Owner_Principal_Name     = $Row.'Owner Principal Name'
                                        Report_Period            = $Row.'Report Period'
                                        department               = $OtherDepartment
                                        ReportPeriod             = $Period
                                        ReportDate               = $RefreshDate
                                        InsertedAt               = (Get-Date)
                                        SourceReport             = $ReportName   
                                        }

                                    Write-Log "Writing $ReportName record into SQLServer ..." -ForegroundColor Cyan
                                    Write-ToSqlTable -TableName $ReportName -Data @($oneDriveObject) 
                                    $RowsInserted++  
                                }
                            } 
                            catch {
                                Write-Log "Error while recovering department for $UserPrincipalName : $_" -ForegroundColor Yellow
                            }
                        }
                        $RowsRetrievedOneDrive = $Data.Count
                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status "SUCCESS" `
                            -RowsRetrieved $RowsRetrievedOneDrive `
                            -RowsInserted $RowsInserted `
                            -DurationSeconds ([int]((Get-Date) - $TaskStart).TotalSeconds)
                    }
                    else {
                        Write-Log "Writing $ReportName records into SQLServer ..." -ForegroundColor Cyan
                        Write-ToSqlTable -TableName $ReportName -Data $Data                   
                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status "SUCCESS" `
                            -RowsRetrieved $Data.Count `
                            -RowsInserted $Data.Count `
                            -DurationSeconds ([int]((Get-Date) - $TaskStart).TotalSeconds)

                        $OthersInserted = $Data.Count
                        $OthersRetrieved = $Data.Count
                    }
                    $TotalRowsRetrieved += $RowsRetrievedExchange
                    $TotalRowsRetrieved += $RowsRetrievedOneDrive
                    $TotalRowsRetrieved += $OthersRetrieved

                    $TotalRowsInserted += $RowsInserted
                    $TotalRowsInserted += $OthersInserted
                }
                else {
                    Write-Log "No data returned for $ReportName" -ForegroundColor Yellow
                }
            }
        } 
        catch {
            Write-Log "Script failed to process $RepotName data: $($_.Exception.Message)" -ForegroundColor Red
            Write-ExecutionLog `
                -ExecutionId $ExecutionId `
                -ReportName $ReportName `
                -RowsRetrieved 0 `
                -RowsInserted 0 `
                -Status "FAILED" `
                -ErrorMessage $_.Exception.Message
            throw
        }
    }
    
    #STEP 2: users
    $AllUsers = @()
    $UsersTable = "Users"
    $Url = "https://graph.microsoft.com/v1.0/users"
    Write-Log "STEP 2 - CASE: $UsersTable" -ForegroundColor Cyan

    do {
        $Response = Invoke-GraphRequest -Url $Url -Headers $UserHeaders
        if ($Response -and $Response.value) {
            $AllUsers += $Response.value
            $Url = $Response.'@odata.nextLink'
        }
        else {
            $Url = $null
        }
    } while ($Url)

    try {
        
        $RowsInserted = 0
        foreach ($User in $AllUsers) {
            Write-Log "Retrieving user $($User.userPrincipalName) details" -ForegroundColor Cyan
            $userObject = [PSCustomObject]@{
                id                = $User.id
                displayName       = $User.displayName
                userPrincipalName = $User.userPrincipalName
                mail              = $User.mail
                department        = $User.department
                jobTitle          = $User.jobTitle
                accountEnabled    = $User.accountEnabled
                createdDateTime   = $User.createdDateTime
            }
            Write-Log "Writing $($User.userPrincipalName) details into SQLServer ..." -ForegroundColor Cyan
            Write-ToSqlTable -TableName $UsersTable -Data $userObject
            $RowsInserted++
        }
        $TotalRowsInsertedUsers = $RowsInserted
        $TotalRowsRetrievedUsers = $AllUsers.Count

        Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName $UsersTable `
            -Status "SUCCESS" `
            -RowsRetrieved $TotalRowsRetrievedUsers `
            -RowsInserted $RowsInserted `
            -DurationSeconds ([int]((Get-Date) - $TaskStart).TotalSeconds)
    }
    catch {
        Write-Log "Script failed to process $UsersTable data: $($_.Exception.Message)" -ForegroundColor Red
        Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName $UsersTable `
            -Status "FAILED" `
            -RowsRetrieved 0 `
            -RowsInserted 0 `
            -ErrorMessage $_.Exception.Message
        throw
    }
    
    Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName "TOTAL" `
            -Status "SUCCESS" `
            -RowsRetrieved ($TotalRowsRetrieved + $TotalRowsRetrievedUsers) `
            -RowsInserted ($TotalRowsInserted + $TotalRowsInsertedUsers) `
            -DurationSeconds ([int]((Get-Date) - $TaskStart).TotalSeconds)

    #STEP 3: powerbi dashboard
    Write-Log "STEP 3 - Creating final metrics for PowerBI" -ForegroundColor Cyan

    $Date =  Get-Date -Format "yyyy-MM-dd"
    $createPowerBIDataModel = 
    "
        WITH Data AS (
            SELECT
                SUM(ISNULL([Primary_Item_Count],0)) AS Exchange_Total_Primary_Item_Count,
                SUM(ISNULL([Archive_Item_Count],0)) AS Exchange_Total_Archive_Item_Count,
                CAST(ROUND(SUM(ISNULL([Primary_Total_Size_Bytes],0)) / 1073741824.0, 2) AS DECIMAL(18,2)) AS Exchange_Total_Primary_Total_Size_GB,
                CAST(ROUND(SUM(ISNULL([Archive_Total_Size_Bytes],0)) / 1073741824.0, 2) AS DECIMAL(18,2)) AS Exchange_Total_Archive_Total_Size_GB,
                SUM(ISNULL(o.[File_Count],0)) AS OneDrive_Total_File_Count,
                SUM(ISNULL(o.[StorageUsedGB],0)) AS OneDrive_Total_StorageUsedGB,
                SUM(ISNULL(s.[File_Count],0)) AS SharePoint_Total_File_Count,
                SUM(ISNULL(s.[StorageUsedGB],0)) AS SharePoint_Total_StorageUsedGB,
                COUNT(DISTINCT u.[UserPrincipalName]) AS Users_Total

            FROM [DataCare].[dbo].[Exchange] e
            CROSS JOIN [DataCare].[dbo].[OneDrive] o
            CROSS JOIN [DataCare].[dbo].[SharePoint] s
            CROSS JOIN [DataCare].[dbo].[Users] u
            WHERE u.[UserPrincipalName] IS NOT NULL
        )
    "
    Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $createPowerBIDataModel
    Write-Log "PowerBI datacare data model created successfully." Green

    $createPowerBITables = "
        -- 1st view on powerbi and backup datacare
        INSERT INTO [DataCare].[dbo].[PowerBIDataModelBackup_$($Date)]
        SELECT * FROM Data;

        -- 2nd view on powerbi
        INSERT INTO [DataCare].[dbo].[PowerBIDataModel]
        SELECT * FROM Data;
    "
    Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $createPowerBITables
    Write-Log "[DataCare].[dbo].[DashboardDataBackup_$($Date)] and DashboardDataModel created successfully." Green
    
    Write-Log "=== END DATACARE EXE ===" -ForegroundColor Green
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName "TOTAL" `
            -Status "FAILED" `
            -RowsRetrieved 0 `
            -RowsInserted 0 `
            -ErrorMessage $_.Exception.Message
    throw
}