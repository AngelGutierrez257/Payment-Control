
## Create database Gestao Financeira (Financial Management)

### original Excel file structure:

  The original Excel file, MapaPagamentos2026, has several sheets that are not relevant for the initial objective. The relevant sheet is named MapaPagamentos2026 as well, and the staging payments
table reflects the structure of this sheet, adding only the LoadDate column that is a derived column defined in a control flow task and populated with an SSIS variable, and the ImportID that is populated on import, as defined by the IDENTITY keyword to auto increment the id number of the imported rows, starting by one, adding one. 

#### 1. MapaPagamentos_Stagging:

```sql
USE [Gestao Financeira]
GO
drop table if exists [dbo].[MapaPagamentos_Stagging];
go


CREATE TABLE [dbo].[MapaPagamentos_Stagging](
	[Ano] [float] NULL,
	[Mes] [float] NULL,
	[Dia] [float] NULL,
	[Data] [datetime] NULL,
	[Dia da semana] [nvarchar](255) NULL,
	[Codigo Fornecedor] [float] NULL,
	[Nome Entidade] [nvarchar](255) NULL,
	[Tipo Entidade] [nvarchar](255) NULL,
	[Tipo Fornecimento] [nvarchar](255) NULL,
	[Empresa] [nvarchar](255) NULL,
	[Descricao] [nvarchar](255) NULL,
	[Forma Pagamento] [nvarchar](255) NULL,
	[Valor] [float] NULL,
	[Banco de debito Previsto] [nvarchar](255) NULL,
	[Obs adicionais] [nvarchar](255) NULL,
	[Entidade beneficiada confirming] [nvarchar](255) NULL,
	[Em atraso?] [nvarchar](255) NULL,
	[Situacao actual] [nvarchar](255) NULL,
	[Data Situação actual] [datetime] NULL,
	[Banco Pagamento] [nvarchar](255) NULL,
	[LoadDate] [date] NULL,
	[ImportID] [int] IDENTITY(1,1) NOT NULL,
  	constraint ImportID primary key ([ImportID]) 
);
go

```




The data types defined also reflect the data types defined in the source Excel sheet, and the names of the columns are the exact column names. 



#### 2. Staging_Fornecedeores (Suppliers)

```sql
USE [Gestao Financeira]
GO

drop table if exists [fornecedores].[Staging_Fornecedores];
go

CREATE TABLE [fornecedores].[Staging_Fornecedores](
	[Cód] [float] NULL,
	[Nome] [nvarchar](500) NULL,
	[Localidade] [nvarchar](500) NULL,
	[Telefone] [nvarchar](255) NULL,
	[Desig.Comercial] [nvarchar](500) NULL,
	[NºContrib] [nvarchar](255) NULL,
	[ImportDate] [date] NULL
) ON [PRIMARY]
GO

--ALTER TABLE [fornecedores].[Staging_Fornecedores] ADD  DEFAULT (CONVERT([date],getdate())) FOR [ImportDate]
--GO
```

The alter table was for testing. The ssis etl pipeline provides the ImportDate data through a derived column transformation.



#### 3. Dim Calendar 

To create and populate the Date table, mssqltips has an excellent CTE approach that defines local settings such as the start of the week, starts a sequence of dates and from that initial sequence all other date calculations can be determined using basic date functions. You can find the explanation and original code here: https://www.mssqltips.com/sqlservertip/4054/creating-a-date-dimension-or-calendar-table-in-sql-server/ 

```sql

Use [Gestao Financeira] 
go

drop table if exists [dbo].[Calendario];
go

set datefirst 1,
	dateformat mdy,
	language US_ENGLISH;
```

In this first block, the first day of the week is defined. As this is for a portuguese company, the first day is Monday (1). The other parameters were left as the original.

```sql
declare @StartDate date = '20100101';
declare @CutoffDate date = dateadd(DAY, -1, dateadd(YEAR,40, @StartDate));

```

After, we declare the initial date, respecting the date format defined above, and declare the final date as the day before the 01 of January, adding 40 years to the start date with dateadd(datepart, the number of periods to add to that datepart and finally a value that resolves to a datetime. In this case, the startdate, as it also accepts variables)

Then, the center of the Dim Calendar definition is composed of recursive cte's:

```sql
;with seq(n) as
(
	select 0 union all select n + 1 from seq
	where n < datediff(day, @StartDate, @CutoffDate)
	),
	d(d) as 
	(
	select dateadd(day, n, @StartDate) from seq
	),
	src as 
	(
	select
	[Data]		= convert(date,d),
	Dia			= datepart(day, d),
	NomeDia		= datename(weekday, d),
	Semana		= datepart(week, d),
	SemanaISO	= datepart(iso_week, d),
	DiadaSemana	= datepart(weekday, d),
	Mes			= datepart(month, d),
	NomeMes		= datename(month, d),
	Trimestre	= datepart(quarter, d),
	Ano			= datepart(year, d),
	PrimDiaMes	= datefromparts(year(d), month(d), 1),
	UltDiaAno	= datefromparts(year(d), 12, 31),
	DiadoAno	= datepart(dayofyear, d),
	TrimestreFiscal = case 
						when month(d) in (1,2,3) then 1
						when month(d) in (4,5,6) then 2
						when month(d) in (7,8,9) then 3
					else 4 end
	from d
	)
```

The first cte generates a sequence of number, initiated by select 0. The union all combines all results from the recursive cte defined by select n + 1 from seq. This is a loop that adds a number to that initial 0. where n < datediff(day, @StartDate, @CutoffDate) calculates the number of days (datediff) and where < n is the stopping condition where the sequence stops when it reaches the number of days that exist between the two days expressed in datediff.

```sql
d(d) as 
(
select dateadd(day, n, @StartDate) from seq
),
```
**d** adds a day for each value in the sequence in n, starting from startdate.

The final cte, **src**, simply applies basic date functions to add date dimensions that can be used in business calculations, starting by converting the string from n in d to date data type. From there, any date function and date transformation can be applied.

Finnaly, the data in the **src** cte is used to create the table dbo.Calendario with a simple select *  into. maxrecursion removes the limit of 100 times a recursive cte is allowed to loop to generate the sequence, as it loops to generate each day. This allows to generate 14 610 dates, or 40 years. The clustered index organizes the data by Date and improves query performance.

```sql
select * into dbo.Calendario
from src
order by [Data]
option (maxrecursion 0);

create unique clustered index PK_DimDate on dbo.Calendario([Data]); 
```

### 4. Dim Empresa (Dim_Company)

This dimension table will contain only two rows, has there are two different companies managed by the same team. This will allow analysis of costs per company, which existed only at a very high level and for accounting proposes. 

```sql

use [Gestao Financeira];
go


drop table if exists dbo.Empresa;
go

create table dbo.Empresa (
ID int identity(1,1) primary key not null,
[Empresa] nvarchar(256) not null,
NIF nvarchar(500) null,
DataInicio date not null,
DataFim date not null,
IsCurrent bit not null,
HashValue varbinary(32) not null
)
```





### 5. Dim Fornecedores Stagging (Import table for suppliers info)

This table is the import/stagging table to receive supplier data coming directly from ssis. This table is not imported from the same excel source as the others tables, but comes, as you can see in the ssis documentation, from a suppliers table extratected manually from the erp program.

```sql
USE [Gestao Financeira]
GO


DROP TABLE IF EXISTS dbo.Staging_Fornecedores;
go

CREATE TABLE dbo.Staging_Fornecedores (
    [Cód]          FLOAT,
    Nome           NVARCHAR(500),
    Localidade     NVARCHAR(500),
    Telefone       NVARCHAR(255),
    [Desig.Comercial] NVARCHAR(500),
    [NºContrib]    NVARCHAR(255),
    ImportDate     DATE DEFAULT CAST(GETDATE() AS DATE)
);

```




### 6. Dim Fornecedores (Dim_Suppliers)

This dimension contains the relevant cleaned information through a merge transformation from the previous stagging table. Similarly to Dim Calendar, the dim tables contain an index to increase query performance and refential integrity.

```sql
USE [Gestao Financeira]
GO

/****** Object:  Table [dbo].[Fornecedores]    Script Date: 05/11/2026 11:47:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Fornecedores](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Codigo] [float] NULL,
	[Nome] [nvarchar](500) NOT NULL,
	[Localidade] [nvarchar](500) NULL,
	[Telefone] [nvarchar](255) NULL,
	[Desig.Comercial] [nvarchar](500) NULL,
	[Nº Contribuinte] [nvarchar](255) NULL,
	[DataInicio] [date] NOT NULL,
	[DataFim] [date] NULL,
	[IsCurrent] [bit] NOT NULL,
	[HashValue] [varbinary](32) NOT NULL,
PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
GO

```

### 7. Dim Tipo Entidade (Entity Type)

This dimension table contains the type of entities defined in the original excel document as well their change over time (SCD). As well, contain an index to increase query performance and referential integrity.

```sql
use [Gestao Financeira]
go

-- Confirm you are in the right context:

 -- select DB_NAME() as currentDatabase;


 drop table if exists dbo.[Tipo Entidade];
 go

 create table dbo.[Tipo Entidade] (
	ID int identity (1,1) primary key not null,
	[Tipo Entidade] nvarchar(256) not null,
	[Descricao] nvarchar(500) not null,
	[DataInicio] date not null,
	[DataFim] date null,
	[IsCurrent] bit not null,
	[HashValue] varbinary(32) not null

 )

	CREATE NONCLUSTERED INDEX IX_TipoEntidade
	ON dbo.[Tipo Entidade]([Tipo Entidade]);

```

### 8. Dim Tipo Fonecimento (Type of Supply)

This dimension table contains the categorizations of the type of supply paid. The primary key provides both better query performance and referential integrity. 

```sql
Use [Gestao Financeira] 
go

drop table if exists [dbo].[Tipo Fornecimento];
go
create table [Tipo Fornecimento] (
ID int identity(1,1) not null primary key,
[Tipo Fornecimento] nvarchar(500) not null,
[Data de Inicio] datetime not null,
[Data de Fim] datetime,
[IsCurrent] bit default 1,
[Hashbytes] varbinary (32)
)
```

### 9. Dim Tipo Pagamento  (Type of payment method)

This is the last dimension table for Payment Control. It contains the payment method used to process a given payment. This information is important for control of payments, so, when a payment made with physical money took place.

```sql
USE [Gestao Financeira]
GO

-- Confirm you are in the right context:

-- select DB_NAME() as currentDatabase;

DROP TABLE IF EXISTS [dbo].[Tipos Pagamento]
GO
CREATE TABLE dbo.[Tipos Pagamento] (
    ID                  INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
    [Tipo de Pagamento]	NVARCHAR(500) NOT NULL,
    Descricao           NVARCHAR(500) NULL,
    
    DataInicio          DATE NOT NULL,
    DataFim             DATE NULL,           
    IsCurrent           BIT NOT NULL DEFAULT 1,
    HashValue           VARBINARY(32) NOT NULL,
    
    INDEX IX_Dim_Clientes_Current (IsCurrent)
);
```
### 10. Fact Pagamentos

This is the main fact table, that contains the cleaned and normalized data inserted from the MapaPagamentos_Stagging table, the dimensional keys that are connected to the dimensions, extracted from the original MapaPagamentos excel table and Fornecedores table, by the respective keys. The index created are the ones based on the keys most used in joins for the views that provide the calculation for the html email sent daily.

```sql
use [Gestao Financeira];
go



 drop table if exists dbo.Pagamentos;
 go

 create table dbo.Pagamentos (
 ID_Pagamento int identity (1,1) primary key not null,
 ImportID int not null, -- Business key
 [Data] date not null, -- all null dates in stagging: in all null dates the row is all null
 BancoID int not null, -- conecta com Banco Pagamento
 FornecedorID int not null,
 [Tipo Fornecimento ID] int not null,
 [Tipo Pagamento ID] int not null,
 [Tipo Entidade ID] int not null,
 EmpresaID int not null,
[Nome Entidade] nvarchar(256) null,
Descricao nvarchar(500) null,
Valor float null,
[Banco de debito Previsto] nvarchar(256)  null,
[Estado Pagamento] nvarchar(256) null, --substitui Em atraso?
[Data Situação actual] date null,
[LoadDate] date not null,
HashValue varbinary(32)
)

CREATE NONCLUSTERED INDEX IX_FactPagamentos_FornecedorID
ON dbo.Pagamentos (FornecedorID);

CREATE NONCLUSTERED INDEX IX_FactPagamentos_PagamentoID
ON dbo.Pagamentos ([Tipo Pagamento ID]);

CREATE NONCLUSTERED INDEX IX_FactPagamentos_TipoFornecimentoID
ON dbo.Pagamentos ([Tipo Fornecimento ID]);

CREATE NONCLUSTERED INDEX IX_FactPagamentos_BancosID
ON dbo.Pagamentos (BancoID);

CREATE NONCLUSTERED INDEX IX_FactPagamentos_Data
ON dbo.Pagamentos (Data);



```
