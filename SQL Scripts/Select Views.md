
## Select Statements on Views

This file contains the sql statements used in the Execute SQL Tasks in the Refresh Views Sequence Container. They are simple SELECTs and self explanatory. 
More details you can get in the Create Views file and in the SSIS Package.md file.

### Select View Pagamentos Atrasados Agregado (Overdue Payments Aggregated):

```sql
SELECT [Year]
      ,[Tipo Entidade]
      ,[Tipo Fornecimento]
      ,[Tipo de pagamento]
      ,[Banco]
      ,[Numero total de pagamentos]
      ,[Montante Total]
  FROM [Gestao Financeira].[dbo].[Pagamentos_Atrasados_Agregado]
order by [Year],[Tipo Entidade]


```


### Select View Pagamentos Due7 Days Agregado (Payments in 7 days agregated):

```sql
SELECT [Year]
      ,[Numero de Semana]
      ,[Data]
      ,[Tipo Entidade]
      ,[Tipo Fornecimento]
      ,[Fornecedor]
      ,[Numero total de pagamentos]
      ,[Montante Total]
      ,[Empresa]
  FROM [Gestao Financeira].[dbo].[Pagamentos_Due7Days_Agregado]
order by [Numero de Semana]

```

### Select View Lista Pagamentos em Atraso (Ovedue payments list):

```sql
SELECT [Data de pagamento]
      ,[Montante]
      ,[Estado Pagamento]
      ,[UltimaActualizacao]
      ,[Nome Entidade]
      ,[Tipo Entidade]
      ,[Tipo Fornecimento]
      ,[Banco de Pagamento]
      ,[Empresa]
  FROM [Gestao Financeira].[dbo].[Pagamentos_em_Atraso]
order by [Data de pagamento]
```

### Select View Payments Due 15 Days:

```sql
SELECT [Data de pagamento]
      ,[Montante]
      ,[Estado Pagamento]
      ,[UltimaActualizacao]
      ,[Código Fornecedor]
      ,[Nome Entidade]
      ,[Tipo Fornecimento]
      ,[Banco de Pagamento]
      ,[Empresa]
  FROM [Gestao Financeira].[dbo].[Payments_Due_15Days]
order by  [Data de pagamento]
```

### Select View Resumo Pagamentos Company ID = 1 (Payments Summary for Company ID = 1):

```sql
SELECT [Numero de pagamentos]
      ,[Montante Total de pagamentos efectuados à data]
      ,[Montante Total Próximos 7 dias]
      ,[Montante Total para a semana seguinte]
      ,[Montante Total Próximos 15 dias]
      ,[Montante total de pagamentos em atraso]
      ,[Montante de fornecimento de material de colocacao em atraso]
      ,[Montante de fornecimento de vidro em atraso]
      ,[Montante de prestações bancárias e leasing em atraso]
      ,[Montante de consumos básicos em atraso]
      ,[Montante de TSU em atraso]
      ,[Montante de IRS em atraso]
      ,[Montante de IUC em atraso]
      ,[Montante de pagamentos com viaturas em atraso]
  FROM [Gestao Financeira].[dbo].[Resumo Pagamentos CrystalCompass]

```

### Select View Resumo Pagamentos Company ID = 2 (Payments Summary Company ID = 2):

```sql
SELECT
       [Montante Total de pagamentos efectuados à data]
      ,[Montante Total Próximos 7 dias]
      ,[Montante Total para a semana seguinte]
      ,[Montante Total Próximos 15 dias]
      ,[Montante total de pagamentos em atraso]
      ,[Montante de prestações bancárias e leasing em atraso]
      ,[Montante de consumos básicos em atraso]
      ,[Montante de TSU em atraso]
      ,[Montante de IRS em atraso]
      ,[Montante de IUC em atraso]
      ,[Montante de pagamentos com viaturas em atraso]
  FROM [Gestao Financeira].[dbo].[Resumo Pagamentos Instalpro]
```

