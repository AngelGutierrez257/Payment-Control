# Gestão Financeira (Payment Control DWH & ETL)

An enterprise-grade Data Warehouse and ETL pipeline engineered with **Microsoft SQL Server (T-SQL)** and **SSIS (SQL Server Integration Services)**. Designed to automate financial payment tracking, eliminate manual Excel errors, and deliver automated executive reporting for multi-company operations.

---

## Architecture & Data Flow

The architecture follows strict Kimball dimensional modeling principles, ingesting raw operational sheets, normalizing dirty data through modular cleansing layers, maintaining historical changes via Slowly Changing Dimensions (SCD Type 2), and exposing optimized views for automated HTML executive briefing generation.

```text
[ Excel / ERP Sources ] 
       │
       ▼ (Win32 P/Invoke File Copy & SSIS Staging)
[ Staging Schema ] (Data Cleansing & Normalization via T-SQL Functions)
       │
       ├─────────────────────────────────────────────────┐
       ▼                                                 ▼
[ Dimension Tables ]                             [ Fact Table (Pagamentos) ]
 (SCD Type 2 & SHA2-256 Hashes)                   (Dual-Hash Change Detection)
       │                                                 │
       └────────────────────────┬────────────────────────┘
                                ▼
               [ Executive Reporting & BI Views ]
                                ▼
         [ Stored Procedure: GenerateFinancialHtmlReport ]
                                ▼
           [ SMTP Dispatch via C# SSIS Script Task ]
```
Key Technical Features
1. Robust Data Cleansing & Normalization
Built modular scalar T-SQL cleansing functions (fn_ClassificacaoBancos, fn_ClassificacaoTipoFornecimento, fn_ClassificacaoTipoEntidade, etc.) to standardize inconsistent string inputs, typos, and accents across manual user entries.

Reusing these functions inside both dimension merges and fact staging guarantees zero data drift between lookups.

2. Advanced SCD Type 2 & Dual-Hash Fact Strategy
Dimensions: Tracks historical changes across suppliers, entity types, and companies using surrogate keys, DataInicio, DataFim, IsCurrent flags, and SHA2-256 change detection hashes.

Fact Table (Pagamentos): Implements a Dual-Hash Strategy separating core transaction identities (BusinessHashValue) from mutable operational states (ControlChangeRowHash), allowing safe updates to payment statuses over time without duplicating records.

3. Operational Resilience (Win32 File-Lock Mitigation)
Solves the common Excel file-lock problem in shared network environments. A custom C# SSIS Script Task utilizes kernel32.dll's CreateFile with non-blocking read/write/delete sharing flags to safely clone operational spreadsheets into a shadow workspace before ingestion.

4. Dynamic HTML Executive Reporting
Developed a cursorless T-SQL stored procedure (dbo.GenerateFinancialHtmlReport) that aggregates metrics from reporting views, leverages FOR XML PATH('tr') and CROSS APPLY (VALUES ...) to pivot summary cards vertically, and formats financial figures dynamically using Portuguese locale settings (pt-PT).

Automatically compiles a responsive HTML document passed into a secure C# SMTP script task (using TLS encryption and environment-protected app passwords) to dispatch daily briefing emails.

5. Centralized Event-Driven Error Auditing
Built an OnError package event handler connected to a custom ETL_ErrorLog table via a stored procedure (dbo.ErrorCapture). Captures execution GUIDs (RunID), exact error codes, messages, machine names, and host user states for centralized troubleshooting.

 
Database Engine: Microsoft SQL Server (T-SQL)

ETL Orchestration: SSIS (SQL Server Integration Services) & SQL Server Agent

Design Pattern: Kimball Star Schema / Dimensional DWH

Scripting / Automation: C# (.NET), Win32 API, T-SQL Stored Procedures & Functions

Security: SHA2-256 Hashing, Sensitive Package Parameters, Referential Integrity Constraints

Gestao-Financeira/
│
├── Schema/
│   ├── 01_Staging_Tables.sql            # Raw ingestion staging tables
│   ├── 02_Dimensions.sql                # Dimension tables (Calendar, Suppliers, etc.)
│   └── 03_Fact_Pagamentos.sql           # Core transactional fact table
│
├── Functions/
│   └── Cleansing_Functions.sql          # Standardized T-SQL cleansing scalar functions
│
├── ETL_Merges/
│   └── SCD2_Dimension_Merges.sql        # MERGE statements with SHA2-256 change detection
│
├── Views/
│   └── Reporting_Views.sql              # Single-pass aggregation views for email briefings
│
├── Procedures/
│   └── GenerateFinancialHtmlReport.sql  # Dynamic XML-pivoted HTML report generator
│
└── SSIS/
    ├── Package.dtsx                     # Main SSIS ETL pipeline package
    └── EventHandlers.sql                # Centralized ETL error logging mechanism

Deployment & Execution
Database Setup: Execute the DDL scripts to provision the Gestao Financeira database, schemas, staging tables, dimensions, and the fact table.

Deploy Functions & Procedures: Run the scalar cleansing functions and the HTML generation stored procedure.

Configure SSIS Package: Open Package.dtsx in Visual Studio (SSIS extension), update connection managers to point to your SQL Server instance and local file paths, and set package parameters (SendingEmail, ReceivingEmail, appPass).

Schedule Automation: Deploy the package to the SSIS Catalog and configure a SQL Server Agent Job to execute the package on a recurring weekday schedule.

*Author*
Angel Gutierrez

GitHub: @AngelGutierrez257
