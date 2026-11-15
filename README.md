Payment Control – Financial Data Warehouse & ETL

A practical Data Engineering project built with Microsoft SQL Server, T-SQL and SQL Server Integration Services (SSIS).

Payment Control is designed to centralize financial payment information from operational Excel files, standardize inconsistent data, maintain historical records and provide structured reporting for financial control.

The project demonstrates the implementation of a relational Data Warehouse using Kimball dimensional modeling principles, incremental data processing, historical dimension management, data cleansing and automated reporting.

Rather than relying on a single preconfigured ETL package, the repository provides the SQL implementation and the technical documentation needed to understand and recreate the SSIS orchestration in a local environment.

1. Business Context

Financial payment control often involves information distributed across spreadsheets and operational systems.

When multiple companies, suppliers and payment records are involved, this approach can create several problems:

Inconsistent supplier and company information.
Different naming conventions and manually entered values.
Difficulty tracking changes to supplier information over time.
Limited visibility into payment status and outstanding obligations.
Repetitive manual reporting activities.
Difficulties identifying and investigating data processing errors.

Payment Control addresses these problems by introducing a centralized analytical data model and a repeatable ETL process.

The objective is to transform operational payment information into structured, historically consistent and report-ready data.

2. Solution Architecture

The solution follows a dimensional Data Warehouse architecture inspired by Kimball methodology.

Data flow
```text
     OPERATIONAL DATA SOURCES
       Excel / Financial Files
                 |
                 v
       SSIS DATA INGESTION
                 |
                 v
          STAGING LAYER
     Raw data and initial validation
                 |
                 v
       DATA CLEANSING (T-SQL)
   Standardization and normalization
                 |
          +------+------+
          |             |
          v             v
    DIMENSION TABLES  FACT TABLE
       SCD Type 2    Payment Records
          |             |
          +------+------+
                 |
                 v
         REPORTING LAYER
       SQL Views and Aggregations
                 |
                 v
       EXECUTIVE REPORT GENERATION
          HTML / Email
```
Main components
Component	Technology	Responsibility  
Source data	Excel	Operational payment and supplier information  
ETL orchestration	SSIS	File handling, data ingestion and process coordination  
Staging	SQL Server	Temporary storage of source data  
Data cleansing	T-SQL	Standardization and normalization  
Data Warehouse	SQL Server	Dimensions and payment fact table  
Historical tracking	T-SQL / SHA2-256	Change detection and SCD Type 2  
Reporting	SQL Views	Aggregated financial information  
Automation	C# / SSIS	HTML report generation and email dispatch  
Error auditing	SQL Server	ETL execution error logging  
3. Technical Implementation  
3.1 Data staging and cleansing  

Raw operational information is first loaded into staging tables.

The project includes modular T-SQL cleansing functions to standardize inconsistent values, including supplier classifications, bank classifications, entity types and other manually maintained attributes.

Examples include:

Standardizing inconsistent text values.
Handling accents and spelling variations.
Normalizing classification fields.
Preparing source data for dimensional lookups.

The objective is to apply consistent business rules before loading the Data Warehouse.

3.2 Dimensional modeling – SCD Type 2

The Data Warehouse uses surrogate keys and historical dimension management.

Slowly Changing Dimension Type 2 is implemented to preserve changes in relevant business attributes rather than overwriting historical information.

The implementation uses:

Surrogate keys.
Effective start and end dates (DataInicio, DataFim).
Current-record identification (IsCurrent).
SHA2-256 hashes for change detection.

This approach allows historical payment analysis to retain the dimensional context associated with the relevant period.

3.3 Payment fact table and change detection

The central fact table, Pagamentos, stores payment-related transactional information.

A dual-hash strategy separates two different concerns:

BusinessHashValue

Identifies the core business transaction and helps prevent duplicate payment records.

ControlChangeRowHash

Tracks changes in operational attributes, allowing existing payment records to be updated when their status or other mutable information changes.

This approach supports incremental processing without treating every change in payment status as a new transaction.

3.4 SSIS orchestration and file handling

SSIS coordinates the ETL workflow and connects the source files to the SQL Server database.

The implementation includes C# Script Tasks for file operations.

One of the implemented approaches uses the Windows API CreateFile through P/Invoke to access Excel files with specific file-sharing permissions.

The intention is to mitigate common file-locking problems when operational spreadsheets are being accessed by other processes.

The SSIS workflow also coordinates the execution of the SQL transformations and reporting operations.

3.5 Executive reporting

The reporting layer uses SQL views and stored procedures to prepare financial information.

The stored procedure:

dbo.GenerateFinancialHtmlReport

Generates an HTML report using SQL aggregations and XML-based row generation.

The resulting HTML content is passed to a C# SSIS Script Task for email dispatch through SMTP.

The report formatting uses Portuguese (pt-PT) regional conventions for financial figures.

3.6 ETL error auditing

The project includes a centralized error logging approach using an SSIS OnError event handler and the SQL Server stored procedure:

dbo.ErrorCapture

The logging design captures information such as:

Execution identifier (RunID).
Error codes and messages.
Machine name.
Execution user information.

This provides a foundation for investigating failures and tracking ETL execution problems.

4. Repository Structure

The repository contains the SQL implementation organized by functional responsibility.

text
Payment-Control/
│
├── Schema/
│   ├── 01_Staging_Tables.sql
│   ├── 02_Dimensions.sql
│   └── 03_Fact_Pagamentos.sql
│
├── Functions/
│   └── Cleansing_Functions.sql
│
├── ETL_Merges/
│   └── SCD2_Dimension_Merges.sql
│
├── Views/
│   └── Reporting_Views.sql
│
├── Procedures/
│   └── GenerateFinancialHtmlReport.sql
│
├── SSIS/
│   ├── Package.dtsx
│   └── EventHandlers.sql
│
└── README.md


The SQL scripts are the foundation of the database implementation. The SSIS package provides the orchestration layer.

5. Implementation Guide

The project is designed to be recreated in a local SQL Server and Visual Studio environment.

The SSIS package should not be treated as a universally preconfigured, ready-to-run deployment. Connection managers, file paths, credentials and environment-specific settings must be configured locally.

Step 1 – Prepare the environment

Required technologies:

Microsoft SQL Server.
SQL Server Management Studio (SSMS).
Visual Studio with the SQL Server Integration Services Projects extension.
Microsoft Excel Database Engine / Access Database Engine provider compatible with the Excel connections.
SQL Server Agent (for scheduled execution).

C# Script Tasks also require a compatible SSIS development environment.

Step 2 – Create the database

Create the database:

Gestao Financeira

Execute the SQL scripts in the following logical order:

Create staging tables.
Create dimension tables.
Create the payment fact table.
Deploy cleansing functions.
Deploy dimension and fact loading transformations.
Deploy reporting views.
Deploy reporting and error-handling procedures.

Ensure the database objects and dependencies are available before configuring the SSIS workflow.

Step 3 – Prepare the source files

Configure the Excel source files expected by the ETL process.

The file locations and names must match the configuration used in the SSIS workflow.

Update any local file paths and directory variables to reflect your environment.

Step 4 – Recreate the SSIS package

Open Visual Studio and create an SSIS project.

Follow the documented Payment Control ETL workflow to recreate the package.

The implementation involves:

Creating the required package connection managers.
Configuring Excel and SQL Server connections.
Creating the package variables and parameters.
Implementing the file-copy and file-handling tasks.
Configuring the staging data flows.
Connecting the SQL transformation and loading operations.
Configuring the execution sequence and precedence constraints.
Implementing the C# Script Tasks.
Configuring the error event handler.

The connection managers must be recreated or reconfigured to use the local SQL Server instance and the appropriate Excel files.

Important: The included .dtsx file should be considered a project reference rather than a portable deployment artifact. Visual Studio and SSIS version differences, package configuration and local environment dependencies can require package reconstruction. The SQL scripts and implementation description provide the basis for that reconstruction.

Step 5 – Configure reporting and email

Configure the package parameters required for email dispatch:

SendingEmail
ReceivingEmail
appPass

Provide the appropriate SMTP configuration and credentials for your email provider.

Sensitive credentials should be supplied through local configuration or an appropriately protected deployment environment, not committed as plain text to the repository.

Step 6 – Validate execution

Before scheduling the workflow, validate each component independently:

Source file accessibility.
Excel connection managers.
SQL Server connectivity.
Staging data ingestion.
Data cleansing operations.
Dimension loading and historical tracking.
Payment fact table loading.
Reporting view results.
HTML report generation.
Error logging.
Email dispatch.

Once the complete workflow has been tested locally, it can be deployed to the SSIS Catalog and scheduled using SQL Server Agent.

6. Key Engineering Concepts Demonstrated
Area	Concepts
Data Warehousing	Kimball methodology, star schema, dimensional modeling
SQL Development	T-SQL, stored procedures, views, scalar functions
ETL	SSIS, staging, incremental processing
Data Quality	Cleansing, standardization, normalization
Historical Data	SCD Type 2, surrogate keys, effective dates
Change Detection	SHA2-256, business hashes, row-change hashes
Automation	C#, SSIS Script Tasks, SMTP
Error Management	SSIS event handlers, centralized error logging
7. Project Scope and Limitations

Payment Control is a Data Engineering project developed to demonstrate the design and implementation of a financial Data Warehouse and its associated ETL processes.

A few implementation considerations are important:

The solution requires local configuration of SQL Server, Excel connections and file paths.
The SSIS package may need to be recreated or adapted depending on the Visual Studio and SSIS versions installed.
The project is not a universally deployable production system. It is meant to be deployed in local networks of small companies.
The implementation is designed around the supplied SQL scripts and the documented ETL workflow.
Email functionality requires valid SMTP credentials and provider-specific configuration.

The repository focuses on demonstrating the underlying data engineering concepts and implementation decisions rather than providing a one-click deployment.

8. Project Objective

The main objective of Payment Control is to demonstrate how operational financial information can be transformed into a structured analytical Data Warehouse through SQL Server and SSIS and to provide to management, daily, relevant payment information and control.

It brings together database design, data cleansing, historical tracking, incremental data processing, ETL orchestration and automated reporting in one practical project.

The emphasis is on understanding and implementing the complete data flow, from operational source files to structured financial reporting.

Author: Angel Gutierrez

GitHub: @AngelGutierrez257

Repository: Payment-Control
