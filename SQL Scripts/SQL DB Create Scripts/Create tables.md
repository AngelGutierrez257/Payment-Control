
## Create database Gestao Financeira (Financial Management)

### original Excel file structure:

  The original Excel file, MapaPgamanetos2026, has several sheets that are not relevant for the initial objective. The relevant sheet is named MapaPagamentos2026 as well, and the staging payments
table reflects the structure of this sheet, adding only the LoadDate column that is a derived column defined in a control flow task and populated with an SSIS variable. Those will be explained later.

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

The data types defined also reflect the data types defined in the source Excel sheet, and these are the exact column names.
