
## Create Views Scripts

This file contains he views used to calculate the agreggated values used in the html tables for the daily email. The objective here is to provide, besides an abstraction layer, separate the business logic 
from pure data warehousing. The main objective is to provide the business owners and the financial team with a map for the payments dyue for next week, for the next 15 days, 
the  payments that are overdue and a summary payment tabke for each company since the beggining of the year. 

### 1. Payments_Due_7Days View

This view provides the relevant data for the individual payments that are due in the next 7 days. It connects to the relevant dim tables and, through window functions, avoids duplicate values that can be present in the fact tables 
(the original excel source is manually populated; this acts as a last filter) and defines the 7 days with a simple where clause using the GETDATE() function. the where clause in the cte also provides a filter 
only select valid payments (declared value, Data de Pagamento (payment date) is due during that period or becomes overdue. This is specially important for payments that become overdue in the first day). Aditionally, 
the dims are filtered by IsCurrent condition.

```sql 
USE [Gestao Financeira]
GO

CREATE OR ALTER     view [dbo].[Payments_Due_7Days]
 as
 with cte_data as (
 select p.[Data] as [Data de pagamento], p.[Valor] as [Montante], p.[Estado Pagamento], p.[Data Situação actual] as [UltimaActualizacao],
 f.[ID] as [Código Fornecedor], p.[Nome Entidade] as [Nome Entidade], tf.[Tipo Fornecimento] as [Tipo Fornecimento],
 b.[Nome Banco] as [Banco de Pagamento], e.[Empresa] as [Empresa],
 row_number() over (
	partition by p.[Data], p.[Nome Entidade], p.Valor, p.[Estado Pagamento]
	order by p.[Data] 
	) as rownumber
from [dbo].[Pagamentos] as p
 left join [dbo].[Fornecedores] as f on p.[FornecedorID] = f.[ID] and f.[IsCurrent] = 1
 left join [dbo].[Tipo Fornecimento] as tf on p.[Tipo Fornecimento ID] = tf.[ID] and tf.[IsCurrent] = 1
 left join [dbo].[Bancos] as b on p.[BancoID]=b.[ID] and b.[IsCurrent] = 1
 left join [dbo].[Empresa] as e on p.[EmpresaID] = e.[ID] and e.[IsCurrent] = 1
 where p.[Data] between cast(getdate() as date) and dateadd(day,7, cast(getdate() as date)) and p.[Valor] is not null and p.[Estado Pagamento] like '%prev%' or p.[Estado Pagamento] like '%atras%'
 )
 select [Empresa], [Data de pagamento],[UltimaActualizacao],[Código Fornecedor],[Nome Entidade], [Montante],[Tipo Fornecimento], [Estado Pagamento],[Banco de Pagamento]
 from cte_data 
 where rownumber = 1;
GO
```



### 2. Payments_Due_15_Days 

This mantains the same logic as the previous view but retrieves the payments to be made in the period of 7 days next to the first. So, from day 8 to day 15, as defined in the casted getdate as date and 
where the days period is calculated using DATEADD (provides the difference between dates). 

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER   view [dbo].[Payments_Due_15Days]
 as
 with cte_data as (
 select p.[Data] as [Data de pagamento], p.[Valor] as [Montante], p.[Estado Pagamento], p.[Data Situação actual] as [UltimaActualizacao],
 f.[ID] as[Código Fornecedor], p.[Nome Entidade] as [Nome Entidade], tf.[Tipo Fornecimento] as [Tipo Fornecimento],
 b.[Nome Banco] as [Banco de Pagamento], e.[Empresa] as [Empresa],e.ID as [EmpresaID],p.[BusinessHashValue],
 row_number() over (
	partition by p.[BusinessHashValue]
	order by p.[Data] ) as rownumber
 from [dbo].[Pagamentos] as p
 left join [dbo].[Fornecedores] as f on p.[FornecedorID] = f.[ID] and f.[IsCurrent] = 1
 left join [dbo].[Tipo Fornecimento] as tf on p.[Tipo Fornecimento ID] = tf.[ID] and tf.[IsCurrent] = 1
 left join [dbo].[Bancos] as b on p.[BancoID]=b.[ID] and b.[IsCurrent] = 1
 left join [dbo].[Empresa] as e on p.[EmpresaID] = e.[ID] and e.[IsCurrent] = 1
 where p.[Data] >= dateadd(day,8 , cast(getdate() as date)) and p.[Data] <= dateadd(day,15, cast(getdate() as date)) and p.[Valor] is not null and (p.[Estado Pagamento] like  '%Previ%' or p.[Estado Pagamento] like '%atra%')
 )
 select [Data de pagamento],[Empresa],[UltimaActualizacao],[Código Fornecedor], [Nome Entidade], [Montante],[Tipo Fornecimento], [Estado Pagamento],[Banco de Pagamento]
 from cte_data 
 where rownumber = 1;
GO
```

### 3. Pagamentos_Atrasados_Agregado (aggregated) 

This table provides the agreggated values for payments overdue, giving a control table to know the satet of oversue payments and providing management an information tool to manage payments in light of cash availability. 
This is achieved using simple agregation fucntions such as COUNT and SUM (). The rest of the logic is the similar to the previous views.

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER   view [dbo].[Pagamentos_Atrasados_Agregado] as 

select datepart(year, p.[Data]) as [Year], te.[Tipo Entidade] as [Tipo Entidade], 
tf.[Tipo Fornecimento] as [Tipo Fornecimento],pg.[Tipo de Pagamento] as [Tipo de pagamento], b.[Nome Banco] as [Banco], 
count(*) as [Numero total de pagamentos], sum(p.[Valor]) as [Montante Total] 
from [dbo].[Pagamentos] p 
left join [dbo].[Tipo Entidade] as te on p.[Tipo Entidade ID] = te.[ID] and te.[IsCurrent] = 1
left join [dbo].[Tipo Fornecimento] as tf on p.[Tipo Fornecimento ID] = tf.[ID] and tf.[IsCurrent] = 1
left join [dbo].[Bancos] as b on p.[BancoID]=b.[ID] and b.[IsCurrent] = 1
left join [dbo].[Tipo Pagamento] as pg on p.[Tipo Pagamento ID] = pg.ID
where  p.[Estado Pagamento] like '%atras%' 
group by datepart(year, p.[Data]), te.[Tipo Entidade], tf.[Tipo Fornecimento], pg.[Tipo de Pagamento], b.[Nome Banco];
GO



```

### 4. Pagamentos_em_Atraso

This table provides a list of all overdue payments registered. 

```sql
USE [Gestao Financeira]
GO

 CREATE OR ALTER   view [dbo].[Pagamentos_em_Atraso] 
 as
select p.[Data] as [Data de pagamento], isnull(p.[Valor],0) as [Montante], p.[Estado Pagamento], isnull(p.[Data Situação actual], getdate()) as [UltimaActualizacao],
 /*f.[ID] as[Código Fornecedor], f.[Nome] as [Nome Entidade],*/ p.[Nome Entidade] as [Nome Entidade], te.[Tipo Entidade] as [Tipo Entidade], tf.[Tipo Fornecimento] as [Tipo Fornecimento],
 b.[Nome Banco] as [Banco de Pagamento], e.[Empresa] as [Empresa]
 from [dbo].[Pagamentos] as p
 left join [dbo].[Fornecedores] as f on p.[Nome Entidade] = f.[Nome] and f.[IsCurrent] = 1
 left join [dbo].[Tipo Entidade] as te on p.[Tipo Entidade ID] = te.[ID] and te.[IsCurrent] = 1
 left join [dbo].[Tipo Fornecimento] as tf on p.[Tipo Fornecimento ID] = tf.[ID] and tf.[IsCurrent] = 1
 left join [dbo].[Bancos] as b on p.[BancoID]=b.[ID] and b.[IsCurrent] = 1
 left join [dbo].[Empresa] as e on p.[EmpresaID] = e.[ID] and e.[IsCurrent] = 1
 where p.[Estado Pagamento] like '%atraso%' 
GO

```

The two final views are summary tables, that provide agregations for all payments made since the beggining of the year in number count, value, the next payments due and the payments overdue by relevant type of supply. 
Special mention to taxes and social contributions, as the local administration is quite strict in any overdue on payments of such types.

### 5. Resumo_Pagamentos_<CompanyName>

In this view, for maintainability and logic update reasons, each column is calculated through a subquery, providing logic isolation as well as making each column more modular to maintain. There is performance penalties for this approach, but the amount of data makes it negligeble.
Overdue in Portuguese is the word atraso, that you can see filtered in each column: where [Estado Pagamento] like '%atraso%'

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER   view [dbo].[Resumo Pagamentos Company ID = 1]
 as

 select 
	  count(*) as [Numero de pagamentos], 
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] <= cast(getdate()as date) and [EmpresaID] = 1 and ([Estado Pagamento] like '%Reg%' or [Estado Pagamento] like '%reg%')) as [Montante Total de pagamentos efectuados à data], 
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= cast(getdate() as date) and [Data] <= dateadd(day,7, cast(getdate() as date)) and [EmpresaID] = 1) as [Montante Total Próximos 7 dias],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= dateadd(day,8, cast(getdate() as date)) and [Data] <= dateadd(day,15, cast(getdate() as date)) and [EmpresaID] = 1 ) as [Montante Total para a semana seguinte],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= cast(getdate() as date) and [Data] <= dateadd(day,15, cast(getdate() as date)) and [EmpresaID] = 1 ) as [Montante Total Próximos 15 dias],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [EmpresaID] = 1) as [Montante total de pagamentos em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 8 and [EmpresaID] = 1) as [Montante de fornecimento de material de colocacao em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 17 and [EmpresaID] = 1 ) as [Montante de fornecimento de vidro em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 3 and [EmpresaID] = 1) as [Montante de prestações bancárias e leasing em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 1 and [EmpresaID] = 1) as [Montante de consumos básicos em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 15 and [EmpresaID] = 1) as [Montante de TSU em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 5 and [EmpresaID] = 1) as [Montante de IRS em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 6 and [EmpresaID] = 1) as [Montante de IUC em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 16 and [EmpresaID] = 1) as [Montante de pagamentos com viaturas em atraso]
from [dbo].[Pagamentos]
	where [EmpresaID] = 1 and [Data] <= cast(getdate()as date) and ([Estado Pagamento] like '%Reg%' or [Estado Pagamento] like '%reg%');
GO

```

Explaining by column:
- the [Numero de pagamentos] column counts all payments made under the filter in the where clause, so, is the company with id = 1, contains all payments made today or in earlier dates and the payment state is registered has paid.
- The [Montante Total Próximos 7 dias] provides the total amount of payments due next week related to company id = 1
- Column [Montante Total para a semana seguinte] provides similar logic as the previous column but for the week after.
- [Montante Total Próximos 15 dias] provides the total amount for both weeks. 
- [Montante total de pagamentos em atraso] is the amount of payments overdue, in total, for the company with id = 1.
- [Montante de fornecimento de material de colocacao em atraso] provides the amount of payments overdue for this type of supply (material de colocação em atraso). All subsquent columns have similar logic for each relevant type of supply to this particularly company.


### 6. View Resumo Pagamentos Company ID = 2 

This view is very similar to the previous one. It contains only the relevant types of supply to the company ID = 2.

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER     view [dbo].[Resumo Pagamentos Instalpro]
 as

 select 
	  count(*) as [Numero de pagamentos], 
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] <= cast(getdate()as date) and [EmpresaID] = 2 and [Estado Pagamento] like '%Regularizado%') as [Montante Total de pagamentos efectuados à data], 
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= cast(getdate() as date) and [Data] <= dateadd(day,7, cast(getdate() as date)) and [EmpresaID] = 2) as [Montante Total Próximos 7 dias],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= dateadd(day,8, cast(getdate() as date)) and [Data] <= dateadd(day,15, cast(getdate() as date)) and [EmpresaID] = 2 ) as [Montante Total para a semana seguinte],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Data] >= cast(getdate() as date) and [Data] <=  dateadd(day,15, cast(getdate() as date)) and [EmpresaID] = 2 ) as [Montante Total Próximos 15 dias],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [EmpresaID] = 2) as [Montante total de pagamentos em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 3 and [EmpresaID] = 2) as [Montante de prestações bancárias e leasing em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 1 and [EmpresaID] = 2) as [Montante de consumos básicos em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 15 and [EmpresaID] = 2) as [Montante de TSU em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 5 and [EmpresaID] = 2) as [Montante de IRS em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 6 and [EmpresaID] = 2) as [Montante de IUC em atraso],
	 (select isnull(sum([Valor]),0) from [dbo].[Pagamentos] where [Estado Pagamento] like '%atraso%' and [Data] <= cast(getdate() as date) and [Tipo Fornecimento ID] = 16 and [EmpresaID] = 2) as [Montante de pagamentos com viaturas em atraso]
from [dbo].[Pagamentos]
	where  [EmpresaID] = 2;
GO

```



