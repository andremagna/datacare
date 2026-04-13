# Microsoft 365 DataCare – Enterprise Data Collector

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue)
![SQL Server](https://img.shields.io/badge/SQL%20Server-Express-red)
![Microsoft Graph](https://img.shields.io/badge/API-Microsoft%20Graph-green)
![SonarCloud](https://img.shields.io/badge/Code%20Quality-SonarCloud-brightgreen)
![Status](https://img.shields.io/badge/Status-Production%20Ready-brightgreen)

---

## 📌 Overview

**Microsoft 365 DataCare** is an enterprise-grade PowerShell ETL solution that:

* Extracts Microsoft 365 usage data via Microsoft Graph & Exchange Online
* Enriches data with advanced mailbox statistics
* Loads data into SQL Server
* Provides a Power BI-ready data model
* Integrates CI/CD with SonarCloud for code quality

---

## 🏗️ Architecture

```text
 PowerShell ETL Script
            ↓
Microsoft Graph + Exchange Online 
            ↓
      SQL Server DB
            ↓
        Power BI
```

---

## 📁 Repository Structure

```text
datacare/
│
├── README.md
├── .gitattributes
│
├── app/
│   └── datacare-app.ps1
│   
├── sql/
│   ├── dbo.executionlog
│   ├── dbo.microsoftexchange
│   ├── dbo.microsoftonedrive
│   ├── dbo.microsoftsharepoint
│   ├── dbo.microsoftusers
│   ├── dbo.powerbicountryorregion
│   └── dbo.powerbidatamodelhistory
│
├── powerbi/
│   ├── powerbicalendar
│   ├── powerbicountryorregion
│   └── powerbidatamodelhistory
│
├── certificates/
│   ├── graphcert.cer
│   └── graphcert.pfx
│
└── .github/
    └── workflows/
        └── sonar.yml
```

---

## 📊 Data Sources

### Exchange Online

* Mailbox usage
* Primary & archive statistics
* Recoverable items
* System messages

### OneDrive

* Storage usage
* File counts

### SharePoint

* Site usage
* File counts

### Azure AD

* Users
* Departments
* Country/Region

---

## 🗄️ Database Design

### Database

```
DataCare
```

---

### Core Tables

* **MicrosoftExchange**
* **MicrosoftOneDrive**
* **MicrosoftSharePoint**
* **MicrosoftUsers**
* **ExecutionLog**

---

### 📊 PowerBI Tables

* **PowerBIDataModelHistory**
* **PowerBICountryOrRegion**

---

## ⚙️ Configuration

File:

```
app/datacare-app.ps1
```

```powershell
$Config = @{
    TenantId = "<TENANT_ID>"
    ClientId = "<CLIENT_ID>"
    CertificateThumbprint = "<CERT_THUMBPRINT>"

    Sql = @{
        Server      = "localhost\SQLEXPRESS"
        SqlDBTarget = "DataCare"
    }

    Execution = @{
        Period = "D180"
    }
}
```

---

## 🔐 Authentication

Uses **certificate-based OAuth2** (no client secrets).

---

## 🔑 Required Permissions

### Microsoft Graph

* Reports.Read.All
* User.Read.All

### Exchange Online

* Exchange.ManageAsApp

---

## ▶️ Execution

```powershell
cd app
.\datacare-app.ps1
```

---

## 🔄 ETL Workflow

1. Exchange / OneDrive / SharePoint extraction
2. Azure AD users extraction
3. Power BI model generation

---

## ⚡ Performance

* SqlBulkCopy (fast inserts)
* Retry logic (429 / 500 / 503)
* Dynamic schema mapping

---

## 📜 Logging

* File → `datacare-log.log`
* SQL → `dbo.ExecutionLog`

---

# ⚡ CI/CD – SonarCloud Integration

The project includes a **GitHub Actions pipeline** for code quality analysis.

## 📄 Workflow File

```
.github/workflows/sonar.yml
```

## 🔧 Pipeline Features

* Runs on every push and pull request to `main`
* Uses Java 17 (required by SonarCloud)
* Caches Sonar packages for performance
* Scans PowerShell codebase

---

## ⚙️ SonarCloud Configuration

### Required GitHub Secret

```
SONAR_TOKEN
```

👉 Add in:

```
GitHub → Settings → Secrets → Actions
```

---

## 📄 Workflow Definition

```yaml
name: SonarCloud PowerShell Scan

on:
  push:
    branches:
      - main
  pull_request:
    branches:
      - main

jobs:
  sonar:
    runs-on: windows-latest

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Setup Java 17
        uses: actions/setup-java@v3
        with:
          distribution: temurin
          java-version: 17

      - name: Verify Java
        shell: pwsh
        run: |
          Write-Host "JAVA_HOME = $Env:JAVA_HOME"
          java -version

      - name: Cache Sonar packages
        uses: actions/cache@v4
        with:
          path: ~/.sonar/cache
          key: ${{ runner.os }}-sonar

      - name: SonarCloud Scan
        uses: sonarsource/sonarqube-scan-action@v6
        with:
          projectBaseDir: '.'
        env:
          SONAR_TOKEN: ${{ secrets.SONAR_TOKEN }}
```

---

## 📊 Power BI Integration

* Pre-aggregated KPI table
* Department-level analytics
* Time-based snapshots

---

## 🔒 Security

* Certificate authentication
* No secrets in code
* Secure SQL connection

---

## 🚀 Roadmap

* Incremental loads
* Azure deployment
* Data lake integration
* Monitoring dashboards

---

## 📄 License

© 2026 - Andrea Magnaghi
Internal use only

---

## 👨‍💻 Version

```
1.0.0
```
