
## Create database Gestao Financeira (Financial Management)

### original Excel file structure:

  The original Excel file, MapaPgamanetos2026, has several sheets that are not relevant for the initial objective. The relevant sheet is named MapaPagamentos2026 as well, and the staging payments
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

The alter table was for testing. The ssis etl pipeline provides the ImportDate data through a derived column.



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




