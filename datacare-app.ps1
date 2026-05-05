<#
.SYNOPSIS
    Microsoft 365 Data Collector and SQL Server Loader

.DESCRIPTION
    This script authenticates to Microsoft Graph and Exchange Online using application permissions
    and retrieves usage, activity details and custom fields.

    Retrieved data includes:
        - Exchange mailbox usage:
        https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagedetail?view=graph-rest-1.0&tabs=http
        - OneDrive usage:
        https://learn.microsoft.com/en-us/graph/api/reportroot-getonedriveusageaccountdetail?view=graph-rest-1.0&tabs=http
        - SharePoint site usage:
        https://learn.microsoft.com/en-us/graph/api/reportroot-getsharepointsiteusagedetail?view=graph-rest-1.0&tabs=http
        - Azure AD users usage:
        https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http

    All data is:
        - Logged with structured execution tracking
        - Stored into SQL Server tables (auto-created if not existing)
        - Tracked in dbo.ExecutionLog for auditing
        - Enriched with metadata (period, execution date, source report) [Exchange]

    The script implements:
        - OAuth2 client credentials flow
        - Structured logging with timestamp
        - Automatic modules installations

.REQUIREMENTS
    - PowerShell 5.1 or higher
    - Access to:
        * https://login.microsoftonline.com
        * https://graph.microsoft.com
        * reports*.office.com (Graph reports redirect)
    - SQL Server: server localhost\SQLEXPRESS with: enctyption and TCP/IP network protocol "Enabled"
    - Azure AD App Registration (Application permissions) with:
        * Reports.Read.All (Microsoft Graph) - grant admin consent
        * User.Read.All (Microsoft Graph) - grant admin consent

.CONFIGURATION
    The following variables must be configured:
    - Azure AD:
        $TenantId
        $ClientId
        $CertificateThumbprint

.EXECUTION
    1. Open PowerShell with admin privileges
    2. Navigate to the script folder
    3. Set cerficated properties (only on fist launch) 
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
    #TenantId     = "35734bde-3e33-4eb6-8dd2-0c96b30981bf"
    #ClientId     = "40f32fed-ed3e-48a3-9dd5-891cdf0a39de"
    #CertificateThumbprint = "6DEEF374447A8014CFA7D68C3B2691ACE43C6E4D"
    TenantId     = "76ff1baa-3307-46aa-a752-cc3736d8a2b2"
    ClientId     = "dd80738f-6094-43ff-bf26-03fe4e3bc7da"
    CertificateThumbprint = "4EDBFF09D6B70180A0DC56D8647F6FD44AC99C81"
    BathSettings = @{
        BatchSize  = 200
    }
    Sql = @{
        Server      = "localhost\SQLEXPRESS"
        SqlDBMaster = "master"
        SqlDBTarget = "DataCare-Demo"
        CreateTable_Exchange = "
        IF OBJECT_ID('dbo.MicrosoftExchange','U') IS NULL
        CREATE TABLE dbo.MicrosoftExchange (
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
            DeletedItemSizeGB FLOAT,
            Deleted_Item_Quota__Byte_ BIGINT,
            Has_Archive NVARCHAR(50),
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100),
            Department NVARCHAR(50),
            CountryOrRegion NVARCHAR(50)
        );"
        CreateTable_OneDrive = "
        IF OBJECT_ID('dbo.MicrosoftOneDrive','U') IS NULL
        CREATE TABLE dbo.MicrosoftOneDrive (
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
            CountryOrRegion NVARCHAR(50),
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100)
        );"
        CreateTable_SharePoint = "
        IF OBJECT_ID('dbo.MicrosoftSharePoint','U') IS NULL
        CREATE TABLE dbo.MicrosoftSharePoint (
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
            Department NVARCHAR(50),
            CountryOrRegion NVARCHAR(50),
            Root_Web_Template NVARCHAR(100),
            Owner_Principal_Name NVARCHAR(255),
            Report_Period NVARCHAR(50),
            ReportPeriod NVARCHAR(50),
            ReportDate DATETIME2,
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100)
        );"
        CreateTable_Users = "
        IF OBJECT_ID('dbo.MicrosoftUsers','U') IS NULL
        CREATE TABLE dbo.MicrosoftUsers (
            Id NVARCHAR(255),
            DisplayName NVARCHAR(255),
            UserPrincipalName NVARCHAR(255) NOT NULL,
            Mail NVARCHAR(255),
            Department NVARCHAR(255),
            JobTitle NVARCHAR(255),
            AccountEnabled NVARCHAR(50),
            CreatedDateTime NVARCHAR(50),
            InsertedAt DATETIME2,
            SourceReport NVARCHAR(100),
            CountryOrRegion NVARCHAR(50)
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
            TableSizeMB FLOAT,
            DurationTimeJob NVARCHAR(255),
            ErrorMessage NVARCHAR(MAX),
            MachineName NVARCHAR(255),
            PowerShellVersion NVARCHAR(50)
        );"   
        CreateTable_PowerBICountryOrRegion = "
        IF OBJECT_ID('dbo.PowerBICountryOrRegion','U') IS NULL
        CREATE TABLE dbo.PowerBICountryOrRegion (
            Department NVARCHAR(255),
            CountryName NVARCHAR(MAX),
            CountryCount INT
        );"
    }
    GraphExecution = @{
        Period     = "D180"
        GraphToken = $null
        GraphTokenCreatedAt = $null
        GraphTokenLifetimeMinutes = 55
        GraphTokenSkewMinutes = 5
    }
    SystemParameters = @{
        TotalRowsRetrieved = 0
        TotalRowsInserted  = 0
        TotalRowsInsertedUsers = 0
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
    [System.IO.File]::AppendAllText("$PSScriptRoot\datacare-log.log", "$line`r`n")
    Write-Host $line -ForegroundColor $ForegroundColor
}

function Write-ExecutionLog {
    param(
        [guid]$ExecutionId,
        [string]$ReportName,
        [string]$Status,
        [int]$RowsRetrieved = 0,
        [int]$RowsInserted = 0,
        [string]$ErrorMessage = $null,
        [string]$DurationTimeJob,
        [float]$TableSizeMB = $null
    )

    $MachineName = $env:COMPUTERNAME
    $PSVersion   = $PSVersionTable.PSVersion.ToString()
    if ($ReportName) { $ReportName = $ReportName.Replace("'", "''") } else { $ReportName = "" }
    if ($Status)     { $Status     = $Status.Replace("'", "''") } else { $Status = "" }
    if ($ErrorMessage) { $ErrorMessage = $ErrorMessage.Replace("'", "''") } else { $ErrorMessage = $null }
    if ($DurationTimeJob)     { $DurationTimeJob     = $DurationTimeJob.Replace("'", "''") } else { $DurationTimeJob = "" }

    $query = @"
INSERT INTO dbo.ExecutionLog
(ExecutionId, ExecutionDate, ReportName, Status,
 RowsRetrieved, RowsInserted, DurationTimeJob, 
 ErrorMessage, 
 MachineName, PowerShellVersion, TableSizeMB)
VALUES
('$ExecutionId', SYSDATETIME(), '$ReportName', '$Status',
 $RowsRetrieved, $RowsInserted, '$DurationTimeJob',
 $(if($ErrorMessage){"'$ErrorMessage'"}else{"NULL"}),
 '$MachineName', '$PSVersion',  $(if($TableSizeMB){"$TableSizeMB"}else{"NULL"}
 ))
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

function Database-Init {
    Write-Log "Ensuring database $($Config.Sql.SqlDBTarget) exists..." Cyan

    $createDbQuery = @"
IF DB_ID(N'$($Config.Sql.SqlDBTarget)') IS NULL
    CREATE DATABASE [$($Config.Sql.SqlDBTarget)];
"@

    Invoke-Sqlcmd -ConnectionString $masterConnectionString -Query $createDbQuery

    Write-Log "Database verified/created" Green

    $tables = @{
        ExecutionLog = $Config.Sql.CreateTable_ExecutionLog
        MicrosoftUsers        = $Config.Sql.CreateTable_Users
        MicrosoftExchange     = $Config.Sql.CreateTable_Exchange
        MicrosoftOneDrive     = $Config.Sql.CreateTable_OneDrive
        MicrosoftSharePoint   = $Config.Sql.CreateTable_SharePoint
        PowerBICountryOrRegion = $Config.Sql.CreateTable_PowerBICountryOrRegion
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

function DatabaseTable-DropIfNotEmpty {
    $Tables = @{
        MicrosoftExchange   = $Config.Sql.CreateTable_Exchange
        MicrosoftOneDrive   = $Config.Sql.CreateTable_OneDrive
        MicrosoftSharePoint = $Config.Sql.CreateTable_SharePoint
        MicrosoftUsers      = $Config.Sql.CreateTable_Users
        PowerBICountryOrRegion = $Config.Sql.CreateTable_PowerBICountryOrRegion
    }
    foreach ($table in $Tables.GetEnumerator()) {
        $tableKey = $table.Key
        $createQuery = $table.Value
        $tableName = "dbo.$tableKey"

        $count = Get-ReportCountFromDb -Table $tableKey
        if ($count -gt 0) {
            Write-Log "Number of records in $tableName : $count. Dropping and recreating $tableName table..." Yellow
            $dropQuery = "IF OBJECT_ID('$tableName','U') IS NOT NULL DROP TABLE $tableName;"
            Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $dropQuery
            Write-Log "$tableName table dropped successfully." Green

            Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $createQuery
            Write-Log "Table '$tableName' created successfully" Green
        }
    }
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

        $sqlColumns = Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $schemaQuery | Select-Object -ExpandProperty COLUMN_NAME
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
                        $dr[$sqlCol] = if ($bytes) { [math]::Round(($bytes / 1GB),5) } else { [DBNull]::Value }
                    }
                    else {
                        $dr[$sqlCol] = [DBNull]::Value
                    }
                    continue
                }
                if ($normalizedSql -eq "deleteditemsizegb") {
                    if ($propertyMap.ContainsKey("deleteditemsizebyte")) {
                        $bytes = $row.($propertyMap["deleteditemsizebyte"])
                        $dr[$sqlCol] = if ($bytes) { [math]::Round(($bytes / 1GB),5) } else { [DBNull]::Value }
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
        [string]$Table
    )

    try {
        $tableMap = @{
            "MicrosoftExchange"    = "dbo.MicrosoftExchange"
            "MicrosoftOneDrive"    = "dbo.MicrosoftOneDrive"
            "MicrosoftSharePoint"  = "dbo.MicrosoftSharePoint"
            "MicrosoftUsers"       = "dbo.MicrosoftUsers"
            "PowerBICountryOrRegion" = "dbo.PowerBICountryOrRegion"
        }

        if (-not $tableMap.ContainsKey($Table)) {
            throw "Unknown report name: $Table"
        }

        $tableName = $tableMap[$Table]
        $query = "SELECT COUNT(*) AS Total FROM $tableName"
        $result = Invoke-Sqlcmd `
            -ConnectionString $targetConnectionString `
            -Query $query
        return [int]$result.Total
    }
    catch {
        Write-Log "Failed to retrieve $Table count from DB: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Get-TableSizeMB {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TableName,
        [string]$Schema = "dbo"
    )

    try {
        $query = @"
SELECT 
    SUM(a.data_pages) * 8.0 / 1024 AS TotalMB
FROM sys.tables t
JOIN sys.schemas s ON t.schema_id = s.schema_id
JOIN sys.indexes i ON t.object_id = i.object_id
JOIN sys.partitions p ON i.object_id = p.object_id AND i.index_id = p.index_id
JOIN sys.allocation_units a ON p.partition_id = a.container_id
WHERE t.name = '$TableName'
AND s.name = '$Schema'
"@

        $result = Invoke-Sqlcmd `
            -ConnectionString $targetConnectionString `
            -Query $query `
            -ErrorAction Stop
        return [math]::Round($result.TotalMB, 5)
    }
    catch {
        Write-Log "Error retrieving the size for $Schema.$TableName : $_" Red
        return 0
    }
}

# STORED PROCEDURE
function CreatePowerBIDataModelHistory {
    $query = @"
IF OBJECT_ID('$($Config.Sql.SqlDBTarget).dbo.PowerBIDataModelHistory', 'U') IS NULL
BEGIN
    CREATE TABLE [$($Config.Sql.SqlDBTarget)].[dbo].[PowerBIDataModelHistory]
    (
        ExecutionId UNIQUEIDENTIFIER,
        [Date] DATETIME,
        department NVARCHAR(255),
        Exchange_StorageUsedGB DECIMAL(18,2),
        Exchange_Item_Count BIGINT,
        Exchange_Deleted_Item_Count BIGINT,
        Exchange_DeletedItemSizeGB DECIMAL(18,2),
        OneDrive_Total_File_Count BIGINT,
        OneDrive_Total_StorageUsedGB DECIMAL(18,2),
        SharePoint_Total_File_Count BIGINT,
        SharePoint_Total_StorageUsedGB DECIMAL(18,2),
        Users_Total INT
    )
END;

DECLARE @ExecutionId UNIQUEIDENTIFIER = NEWID();

INSERT INTO [$($Config.Sql.SqlDBTarget)].[dbo].[PowerBIDataModelHistory]
SELECT 
    @ExecutionId,
    GETDATE(),
    e.department,
    e.Exchange_StorageUsedGB,
    e.Exchange_Item_Count,
    e.Exchange_Deleted_Item_Count,
    e.Exchange_DeletedItemSizeGB,
    o.OneDrive_Total_File_Count,
    o.OneDrive_Total_StorageUsedGB,
    s.SharePoint_Total_File_Count,
    s.SharePoint_Total_StorageUsedGB,
    u.Users_Total
FROM
(
    SELECT
        ISNULL(department,'Unknown') AS department,
        SUM(ISNULL([StorageUsedGB],0)) AS Exchange_StorageUsedGB,
        SUM(ISNULL([Item_Count],0)) AS Exchange_Item_Count,
        SUM(ISNULL([Deleted_Item_Count],0)) AS Exchange_Deleted_Item_Count,
        SUM(ISNULL([DeletedItemSizeGB],0)) AS Exchange_DeletedItemSizeGB
    FROM [$($Config.Sql.SqlDBTarget)].[dbo].[MicrosoftExchange]
    WHERE CountryOrRegion != 'Russia'
    GROUP BY ISNULL(department,'Unknown')
) e
LEFT JOIN
(
    SELECT
        ISNULL(department,'Unknown') AS department,
        SUM(ISNULL([File_Count],0)) AS OneDrive_Total_File_Count,
        SUM(ISNULL([StorageUsedGB],0)) AS OneDrive_Total_StorageUsedGB
    FROM [$($Config.Sql.SqlDBTarget)].[dbo].[MicrosoftOneDrive]
    WHERE CountryOrRegion != 'Russia'
    GROUP BY ISNULL(department,'Unknown')
) o ON e.department = o.department
LEFT JOIN
(
    SELECT
        ISNULL(department,'Unknown') AS department,
        SUM(ISNULL([File_Count],0)) AS SharePoint_Total_File_Count,
        SUM(ISNULL([StorageUsedGB],0)) AS SharePoint_Total_StorageUsedGB
    FROM [$($Config.Sql.SqlDBTarget)].[dbo].[MicrosoftSharePoint]
    WHERE CountryOrRegion != 'Russia'
    GROUP BY ISNULL(department,'Unknown')
) s ON e.department = s.department
LEFT JOIN
(
    SELECT
        ISNULL(department,'Unknown') AS department,
        COUNT(DISTINCT [UserPrincipalName]) AS Users_Total
    FROM [$($Config.Sql.SqlDBTarget)].[dbo].[MicrosoftUsers]
    WHERE Mail like '%@ferrero.com' AND AccountEnabled = 'True' AND CountryOrRegion != 'Russia'
    GROUP BY ISNULL(department,'Unknown')
) u ON e.department = u.department;
"@

    Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $query
    Write-Log "[$($Config.Sql.SqlDBTarget)].[dbo].[PowerBIDataModelHistory] table created successfully." Green
}

function CreatePowerBIDataModelCountryOrRegion {
    $queryCountryOrRegion = "
        INSERT INTO dbo.PowerBICountryOrRegion (Department, CountryName, CountryCount)
        SELECT 
            ISNULL(Department, 'Unknown') AS Department,
            CountryOrRegion AS CountryName,
            COUNT(*) AS CountryCount
        FROM dbo.MicrosoftUsers
        WHERE CountryOrRegion IS NOT NULL
        GROUP BY ISNULL(Department, 'Unknown'), CountryOrRegion
        ORDER BY Department, CountryCount DESC;
    "
    Invoke-Sqlcmd -ConnectionString $targetConnectionString -Query $queryCountryOrRegion
    Write-Log "[$($Config.Sql.SqlDBTarget)].[dbo].[PowerBICountryOrRegion] table created successfully." Green
}

# ENTRAID
function Get-GraphAccessToken {
    Write-Log "Requesting Microsoft Graph token using certificate..." Cyan

    try {
        $cert = Get-Item "Cert:\CurrentUser\My\$($Config.CertificateThumbprint)"
        if (-not $cert) {
            throw "Certificate not found with thumbprint $($Config.CertificateThumbprint)"
        }
        if (-not $cert.HasPrivateKey) {
            throw "Certificate does not have a private key"
        }

        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10)

        $header = @{
            alg = "RS256"
            typ = "JWT"
            x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+','-').Replace('/','_')
        }

        $payload = @{
            aud = "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token"
            iss = $Config.ClientId
            sub = $Config.ClientId
            jti = [guid]::NewGuid().ToString()
            nbf = [int]$now.ToUnixTimeSeconds()
            exp = [int]$exp.ToUnixTimeSeconds()
        }

        function ConvertTo-Base64Url($bytes) {
            return [Convert]::ToBase64String($bytes).TrimEnd('=').Replace('+','-').Replace('/','_')
        }

        $headerEncoded  = ConvertTo-Base64Url([Text.Encoding]::UTF8.GetBytes(($header | ConvertTo-Json -Compress)))
        $payloadEncoded = ConvertTo-Base64Url([Text.Encoding]::UTF8.GetBytes(($payload | ConvertTo-Json -Compress)))

        $unsignedToken = "$headerEncoded.$payloadEncoded"
        $bytesToSign = [Text.Encoding]::UTF8.GetBytes($unsignedToken)

        $signatureBytes = $null

        try {
            $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
            if ($rsa) {
                $signatureBytes = $rsa.SignData(
                    $bytesToSign,
                    [System.Security.Cryptography.HashAlgorithmName]::SHA256,
                    [System.Security.Cryptography.RSASignaturePadding]::Pkcs1
                )
            }
        }
        catch {
            Write-Log "Modern RSA method failed, fallback to legacy CSP..." Yellow
        }

        if (-not $signatureBytes) {
            $privateKey = $cert.PrivateKey
            if (-not $privateKey) {
                throw "No usable private key found"
            }
            if ($privateKey -is [System.Security.Cryptography.RSACryptoServiceProvider]) {
                $sha256 = New-Object System.Security.Cryptography.SHA256Managed
                $signatureBytes = $privateKey.SignData($bytesToSign, $sha256)
            }
            else {
                throw "Unsupported private key provider: $($privateKey.GetType().FullName)"
            }
        }

        if (-not $signatureBytes) {
            throw "Failed to sign JWT"
        }

        $signatureEncoded = ConvertTo-Base64Url($signatureBytes)
        $clientAssertion = "$unsignedToken.$signatureEncoded"

        $body = @{
            client_id             = $Config.ClientId
            scope                 = "https://graph.microsoft.com/.default"
            grant_type            = "client_credentials"
            client_assertion_type = "urn:ietf:params:oauth:client-assertion-type:jwt-bearer"
            client_assertion      = $clientAssertion
        }

        $uri = "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token"

        $response = Invoke-RestMethod `
            -Uri $uri `
            -Method POST `
            -Body $body `
            -ContentType "application/x-www-form-urlencoded"

        Write-Log "Graph token acquired (certificate auth)" Green
        return $response.access_token
    }
    catch {
        throw "Graph authentication (certificate) failed: $($_.Exception.Message)"
    }
}

function Get-ValidGraphToken {
    $now = Get-Date

    if ($Config.GraphExecution.GraphToken -and $Config.GraphExecution.GraphTokenCreatedAt) {
        $elapsed = ($now - $Config.GraphExecution.GraphTokenCreatedAt).TotalMinutes
        $remaining = $Config.GraphExecution.GraphTokenLifetimeMinutes - $elapsed
        if ($remaining -gt $Config.GraphExecution.GraphTokenSkewMinutes) {

            Write-Log "Using cached Graph token (age: $([int]$elapsed) min, remaining: $([int]$remaining) min)" -ForegroundColor DarkGray
            return $Config.GraphExecution.GraphToken
        }
        Write-Log "Graph token near expiry (remaining: $([int]$remaining) min) → refreshing..." -ForegroundColor Yellow
    }
    else {
        Write-Log "No cached token, requesting new one..." -ForegroundColor Yellow
    }

    $Config.GraphExecution.GraphToken = $null
    $Config.GraphExecution.GraphTokenCreatedAt = $null

    try {
        $token = Get-GraphAccessToken
        if ([string]::IsNullOrEmpty($token)) {
            throw "Empty token returned from Get-GraphAccessToken"
        }

        $Config.GraphExecution.GraphToken = $token
        $Config.GraphExecution.GraphTokenCreatedAt = Get-Date

        Write-Log "Graph token refreshed successfully" -ForegroundColor Green
        return $Config.GraphExecution.GraphToken
    }
    catch {
        Write-Log "FAILED to acquire Graph token: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Invoke-CustomGraphRequest {
    param (
        [string]$Url,
        [hashtable]$Headers
    )

    Write-Log "Calling: $Url" -ForegroundColor Yellow

    try { return Invoke-RestMethod -Uri $Url -Headers $Headers -Method GET }
    catch {
        $statusCode = $_.Exception.Response.StatusCode.value__
        if ($statusCode -eq 401) {
            Write-Log "401 detected, refreshing Graph token..." -ForegroundColor Yellow

            $newToken = Get-ValidGraphToken
            $Headers["Authorization"] = "Bearer $newToken"

            return Invoke-RestMethod -Uri $Url -Headers $Headers -Method GET
        }

        Write-Log "Request failed: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}


# ======================
#           MAIN
# ======================
try {
    Write-Log "=== START DATACARE ETL ==="
    
    $ExecutionId = [guid]::NewGuid()

    Import-RequiredModule -ModuleName "SqlServer"
    if (-not (Test-SqlConnection)) {
        Write-Log "SQL connection failed." -ForegroundColor Red
        throw "SQL connection failed."
    }
    Database-Init
    DatabaseTable-DropIfNotEmpty

    $AccessToken = Get-ValidGraphToken
    $ReportHeaders = @{
        Authorization = "Bearer $AccessToken"
        Accept        = "text/csv"
    }
    $UserHeaders = @{
        Authorization    = "Bearer $AccessToken"
        ConsistencyLevel = "eventual"
    }

    $TaskStart = Get-Date
    #STEP 1: exchange, onedrive and sharepoint
    $Reports = @{
        MicrosoftExchange   = "https://graph.microsoft.com/v1.0/reports/getMailboxUsageDetail(period='$($Config.GraphExecution.Period)')"
        MicrosoftOneDrive   = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='$($Config.GraphExecution.Period)')"
        MicrosoftSharePoint = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='$($Config.GraphExecution.Period)')"
    }
    foreach ($ReportName in $Reports.Keys) {
        Write-Log "
        ***************************************
        STEP 1 - CASE: $ReportName
        ***************************************" -ForegroundColor Magenta
        try {
            $Response = Invoke-CustomGraphRequest -Url $Reports[$ReportName] -Headers $ReportHeaders
            if ($Response) {
                $Data = $Response | ConvertFrom-Csv
                if ($Data -and $Data.Count -gt 0) {
                    $RowsRetrievedExchange = 0
                    $RowsRetrievedOneDrive = 0
                    $RowsRetrievedSharePoint = 0
                    $RowsInserted  = 0

                    if ($ReportName -eq "MicrosoftExchange") {
                        $TaskStartExchange = Get-Date
                        $batch = @()

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
                            $Url = "https://graph.microsoft.com/v1.0/users/"+$EncodedUpn+"?`$select=department,country"

                            try {
                                $Response = Invoke-CustomGraphRequest -Url $Url -Headers $UserHeaders

                                $exchangeObject = [PSCustomObject]@{
                                    displayName                 = $Row.'Display Name'
                                    userPrincipalName           = $Row.'User Principal Name'
                                    mail                        = $Row.'User Principal Name'
                                    department                  = $Response.department
                                    countryOrRegion             = $Response.country
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
                                $batch += $exchangeObject

                                if ($batch.Count -ge $Config.BathSettings.BatchSize) {
                                    Write-Log "Writing batch of $($batch.Count) records into SQLServer ..." -ForegroundColor Cyan
                                    Write-ToSqlTable -TableName $ReportName -Data $batch
                                    $RowsInserted += $batch.Count
                                    $batch = @()
                                }

                            }
                            catch {
                                Write-Log "Error while recovering department for $UserPrincipalName : $_" -ForegroundColor Red
                            }
                        }

                        if ($batch.Count -gt 0) {
                            Write-Log "Writing final batch of $($batch.Count) records into SQLServer ..." -ForegroundColor Cyan
                            Write-ToSqlTable -TableName $ReportName -Data $batch
                            $RowsInserted += $batch.Count
                        }

                        $duration = (Get-Date) - $TaskStart
                        $durationJob = (Get-Date) - $TaskStartExchange
                        $tableSize = Get-TableSizeMB -TableName $ReportName
                        $RowsRetrievedExchange = $Data.Count
                        $status = if($RowsRetrievedExchange -eq $RowsInserted) {"SUCCESS"} else {"PARTIALLY"}
                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status $status `
                            -RowsRetrieved $RowsRetrievedExchange `
                            -RowsInserted $RowsInserted `
                            -DurationTimeJob  $($durationJob.ToString('hh\:mm\:ss')) `
                            -TableSizeMB $tableSize
                    }
                    elseif ($ReportName -eq "MicrosoftOneDrive") {
                        $TaskStartOneDrive = Get-Date

                        $batch = @()

                        foreach ($Row in $Data) {
                            $ReportRefreshDate = ($Row.PSObject.Properties |
                                Where-Object { $_.Name -like "*Report Refresh Date*" }).Name
                            $RefreshDate = $Row.$ReportRefreshDate

                            $UserPrincipalName = $Row.'Owner Principal Name'.Trim() -replace "^[\uFEFF]", ""
                            if ([string]::IsNullOrEmpty($UserPrincipalName)) {
                                Write-Log "UPN is empty" -ForegroundColor Yellow
                            }

                            $EncodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)
                            $Url = "https://graph.microsoft.com/v1.0/users/"+$EncodedUpn+"?`$select=department,country"

                            try {
                                $Response = Invoke-CustomGraphRequest -Url $Url -Headers $UserHeaders

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
                                    department               = $Response.department
                                    countryOrRegion          = $Response.country
                                    ReportPeriod             = $Period
                                    ReportDate               = $RefreshDate
                                    InsertedAt               = (Get-Date)
                                    SourceReport             = $ReportName    
                                }
                                $batch += $oneDriveObject

                                if ($batch.Count -ge $Config.BathSettings.BatchSize) {
                                    Write-Log "Writing batch of $($batch.Count) $ReportName records into SQLServer ..." -ForegroundColor Cyan
                                    Write-ToSqlTable -TableName $ReportName -Data $batch
                                    $RowsInserted += $batch.Count
                                    $batch = @()
                                }
                            } 
                            catch {
                                Write-Log "Error while recovering department for $UserPrincipalName : $_" -ForegroundColor Yellow
                            }
                        }

                        if ($batch.Count -gt 0) {
                            Write-Log "Final flush: writing $($batch.Count) $ReportName records into SQLServer ..." -ForegroundColor Cyan
                            Write-ToSqlTable -TableName $ReportName -Data $batch
                            $RowsInserted += $batch.Count
                        }

                        $duration = (Get-Date) - $TaskStart
                        $durationJob = (Get-Date) - $TaskStartOneDrive
                        $tableSize = Get-TableSizeMB -TableName $ReportName
                        $RowsRetrievedOneDrive = $Data.Count
                        $status = if($RowsRetrievedOneDrive -eq $RowsInserted) {"SUCCESS"} else {"PARTIALLY"}
                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status $status `
                            -RowsRetrieved $RowsRetrievedOneDrive `
                            -RowsInserted $RowsInserted `
                            -DurationTimeJob  $($durationJob.ToString('hh\:mm\:ss')) `
                            -TableSizeMB $tableSize
                    }
                    else {
                        $TaskStartSharePoint = Get-Date

                        $batch = @()

                        foreach ($Row in $Data) {

                            $ReportRefreshDate = ($Row.PSObject.Properties |
                                Where-Object { $_.Name -like "*Report Refresh Date*" }).Name
                            $RefreshDate = $Row.$ReportRefreshDate

                            $UserPrincipalName = $Row.'Owner Principal Name'.Trim() -replace "^[\uFEFF]", ""
                            if ([string]::IsNullOrEmpty($UserPrincipalName)) {
                                Write-Log "UPN is empty" -ForegroundColor Yellow
                            }

                            $UserDepartment = "NULL"

                            $EncodedUpn = [System.Uri]::EscapeDataString($UserPrincipalName)
                            $Url = "https://graph.microsoft.com/v1.0/users/"+$EncodedUpn+"?`$select=department,country"

                            try {
                                $Response = Invoke-CustomGraphRequest -Url $Url -Headers $UserHeaders
                                $UserDepartment = $Response.department
                            } 
                            catch {
                                Write-Log "Error while recovering department for $UserPrincipalName : $_" -ForegroundColor Yellow
                            }

                            $sharePointObject = [PSCustomObject]@{
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
                                Owner_Principal_Name     = $UserPrincipalName
                                Report_Period            = $Row.'Report Period'
                                Page_View_Count          = $Row.'Page View Count'
                                Visited_Page_Count       = $Row.'Visited Page Count'
                                Root_Web_Template        = $RootWebTemplate
                                department               = $UserDepartment
                                countryOrRegion          = $Response.country
                                ReportPeriod             = $Period
                                ReportDate               = $RefreshDate
                                InsertedAt               = (Get-Date)
                                SourceReport             = $ReportName    
                            }
                            $batch += $sharePointObject

                            if ($batch.Count -ge $Config.BathSettings.BatchSize) {
                                Write-Log "Writing batch of $($batch.Count) $ReportName records into SQLServer ..." -ForegroundColor Cyan
                                Write-ToSqlTable -TableName $ReportName -Data $batch
                                $RowsInserted += $batch.Count
                                $batch = @()
                            }
                        }

                        if ($batch.Count -gt 0) {
                            Write-Log "Final flush: writing $($batch.Count) $ReportName records into SQLServer ..." -ForegroundColor Cyan
                            Write-ToSqlTable -TableName $ReportName -Data $batch
                            $RowsInserted += $batch.Count
                        }

                        $duration = (Get-Date) - $TaskStart
                        $durationJob = (Get-Date) - $TaskStartSharePoint
                        $tableSize = Get-TableSizeMB -TableName $ReportName
                        $RowsRetrievedSharePoint = $Data.Count
                        $status = if($RowsRetrievedSharePoint -eq $RowsInserted) {"SUCCESS"} else {"PARTIALLY"}
                        Write-ExecutionLog `
                            -ExecutionId $ExecutionId `
                            -ReportName $ReportName `
                            -Status $status `
                            -RowsRetrieved $RowsRetrievedSharePoint `
                            -RowsInserted $RowsInserted `
                            -DurationTimeJob  $($durationJob.ToString('hh\:mm\:ss')) `
                            -TableSizeMB $tableSize
                    }

                    $Config.Sql.TotalRowsRetrieved += $RowsRetrievedExchange
                    $Config.Sql.TotalRowsRetrieved += $RowsRetrievedOneDrive
                    $Config.Sql.TotalRowsRetrieved += $RowsRetrievedSharePoint

                    $Config.Sql.TotalRowsInserted += $RowsInserted
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
                -RowsRetrieved $Config.Sql.TotalRowsRetrieved `
                -RowsInserted $Config.Sql.TotalRowsInserted `
                -Status "FAILED" `
                -ErrorMessage $_.Exception.Message
            throw
        }
    }
    
    #STEP 2: users
    Write-Log " 
    ***************************************
    STEP 2 - CASE: Users
    *************************************** " -ForegroundColor Magenta
    $TaskStartUsers = Get-Date
    $AccessToken = Get-ValidGraphToken
    $UserHeaders = @{
        Authorization    = "Bearer $AccessToken"
        ConsistencyLevel = "eventual"
    }

    $AllUsers = @()
    $UsersTable = "MicrosoftUsers"
    $Url = "https://graph.microsoft.com/v1.0/users?`$select=id,displayName,userPrincipalName,mail,department,jobTitle,accountEnabled,createdDateTime,country"

    do {
        $Response = Invoke-CustomGraphRequest -Url $Url -Headers $UserHeaders
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
        $batch = @()

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
                countryOrRegion   = $User.country
            }
            $batch += $userObject

            if ($batch.Count -ge $Config.BathSettings.BatchSize) {
                Write-Log "Writing batch of $($batch.Count) users into SQLServer ..." -ForegroundColor Cyan
                Write-ToSqlTable -TableName $UsersTable -Data $batch
                $RowsInserted += $batch.Count
                $batch = @()
            }
        }

        if ($batch.Count -gt 0) {
            Write-Log "Final flush: writing $($batch.Count) users into SQLServer ..." -ForegroundColor Cyan
            Write-ToSqlTable -TableName $UsersTable -Data $batch
            $RowsInserted += $batch.Count
        }

        $duration = (Get-Date) - $TaskStart
        $durationJob = (Get-Date) - $TaskStartUsers
        $tableSize = Get-TableSizeMB -TableName $ReportName
        $Config.Sql.TotalRowsInsertedUsers = $RowsInserted
        $TotalRowsRetrievedUsers = $AllUsers.Count
        $status = if($TotalRowsRetrievedUsers -eq $Config.Sql.TotalRowsInsertedUsers) {"SUCCESS"} else {"PARTIALLY"}
        Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName $UsersTable `
            -Status $status `
            -RowsRetrieved $TotalRowsRetrievedUsers `
            -RowsInserted $RowsInserted `
            -DurationTimeJob  $($durationJob.ToString('hh\:mm\:ss')) `
            -TableSizeMB $UsersTable
    }
    catch {
        Write-Log "Script failed to process $UsersTable data: $($_.Exception.Message)" -ForegroundColor Red
        Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName $UsersTable `
            -Status "FAILED" `
            -RowsRetrieved $TotalRowsRetrievedUsers `
            -RowsInserted $RowsInserted `
            -ErrorMessage $_.Exception.Message
        throw
    }

    $duration = (Get-Date) - $TaskStart
    $totalRetrieved = ($Config.Sql.TotalRowsRetrieved + $TotalRowsRetrievedUsers)
    $totalInserted = ($Config.Sql.TotalRowsInserted + $Config.Sql.TotalRowsInsertedUsers)
    $status = if($totalRetrieved -eq $totalInserted) {"SUCCESS"} else {"PARTIALLY"}
    Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName "TOTAL" `
            -Status $status `
            -RowsRetrieved $totalRetrieved `
            -RowsInserted $totalRetrieved `
            -DurationTimeJob  $($duration.ToString('hh\:mm\:ss'))

    Write-Log "
    ***************************************
    STEP 3 - Creating final tables for PowerBI
    ***************************************" -ForegroundColor Magenta
    Write-Log "Creating data models for PowerBI ..." -ForegroundColor Cyan
    CreatePowerBIDataModelHistory
    CreatePowerBIDataModelCountryOrRegion
    Write-Log "Data models created successfully" -ForegroundColor Green
    
    Write-Log "=== END DATACARE ETL ===" -ForegroundColor Green
}
catch {
    Write-Log "SCRIPT FAILED: $($_.Exception.Message)" -ForegroundColor Red
    Write-ExecutionLog `
            -ExecutionId $ExecutionId `
            -ReportName "TOTAL" `
            -Status "FAILED" `
            -RowsRetrieved ($Config.Sql.TotalRowsRetrieved + $TotalRowsRetrievedUsers) `
            -RowsInserted ($Config.Sql.TotalRowsInserted + $Config.Sql.TotalRowsInsertedUsers) `
            -ErrorMessage $_.Exception.Message
    throw
}