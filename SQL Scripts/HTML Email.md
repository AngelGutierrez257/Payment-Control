
## HTML Report Store Procedure

The following code gathers the relevant information from the updated views and creates a store procedure that outputs a single value in HTML, which allows to configure the email visuals.

The code is explained block by block, to explain which difficulties were overcome by using that particular statement.

Firstly, defining the procedure and declare the variable. As the length of the email is unknown, as is the characters, the variable is defined as @OutputHTML NVARCHAR(MAX) OUTPUT.

```sql
USE [Gestao Financeira]
GO


CREATE OR ALTER     PROCEDURE [dbo].[GenerateFinancialHtmlReport]
    @OutputHTML NVARCHAR(MAX) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
```


As the tests emails continued to have truncation problems, a solution that worked was to force nvarchar(max) on the variable:

```sql
 SET  @OutputHTML = cast('' as nvarchar(max));
```

The next section was to had to the variable the header styling and content, including the title and calling a GETDATE function to add the date the email is generated and the specific date collation, in this case the ODBC format, as I find it reads better. Tried 103, which is the one who applies to Portugal, but do not fit as well:

```sql
SET @OutputHTML = @OutputHTML +  N'    <html>    <head>    <style>        body { font-family: "Segoe UI", Calibri, Arial, sans-serif; color: #333333; margin: 20px; }        h1 { color: #1a365d; font-size: 24px; margin-bottom: 5px; }        h2 { color: #2b6cb0; font-size: 16px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; margin-top: 30px; margin-bottom: 12px; }        table { border-collapse: collapse; width: 100%; margin-bottom: 10px; font-size: 13px; }        th { background-color: #2b6cb0; color: white; padding: 10px 12px; font-weight: 600; text-align: left; }        td { border: 1px solid #e2e8f0; padding: 10px 12px; text-align: left; }        tr:nth-child(even) { background-color: #f8fafc; }    </style>    </head>    <body>    <h1>Controlo de Pagamentos</h1>    <p style="color: #64748b; font-size: 13px; margin-bottom: 20px;">Gerado automaticamente em: ' + CONVERT(VARCHAR(20), GETDATE(), 120) + '</p>';
```

The first table is the table resulting from querying the view Pagamentos_Due7Days_Agregado, that provides a list of all payments due in the next 7 days. First, a variable is declare to store the output of the select statement, and when selected is cast to xml, that can be read in html. Its data type is forced to avoid truncation. In this case, each column is firstly declare as td = and uses ISNULL to re-format null handling. The column [Montante Total] was also formatted to present currency in the correct Portuguese format. It was ordered by Data asc to show the payments that were more urgent. It is necessary, when converting to XML, to declare the 'tr' delimiter for the columns. The Sql Server parses better the xml that way. So the xml structure will be <tr><td>ColumnValue</td><td>ColumnValue</td></tr> and this defines a line. YOu can see this formatting also when defining the html added to the @OutputHTML variable.

```sql
DECLARE @Tabela1 NVARCHAR(MAX) = CAST( '' as nvarchar(max));
    SELECT @Tabela1 = cast((
    select
        td = ISNULL(CONVERT(VARCHAR(10), [Data], 105), N''),'', 
        td = ISNULL([Tipo Entidade], N''), '',
        td = ISNULL([Tipo Fornecimento], N''), '',
        td = ISNULL([Empresa], N''), '',
        td = ISNULL([Fornecedor], N''), '',
        td = ISNULL(FORMAT([Montante Total], 'C', 'pt-PT'), N'0,00 €'), ''
    FROM dbo.Pagamentos_Due7Days_Agregado
    ORDER BY [Data] ASC
    FOR XML PATH('tr'), TYPE
) AS NVARCHAR(MAX));
```

Then, the resulting xml is added to the @OutputHTML variable, with explicit statement of the column names.  

```sql
    SET @OutputHTML = @OutputHTML + N'<h2>1. Pagamentos Previstos para os Próximos 7 Dias</h2>    <table>        <tr>            <th>Data de Pagamento</th><th>Tipo Entidade</th><th>Tipo Fornecimento</th><th>Empresa</th><th>Fornecedor</th><th>Montante Total</th>        </tr>' + ISNULL(@Tabela1, N'<tr><td colspan="6" style="color: #64748b; font-style: italic;">Nenhum registo encontrado.</td></tr>') + N'</table>';

```

Following the same process, the payments for the next 8th to 15th days were retrieved from the view Payments_Due_15Days:

```sql
DECLARE @Tabela2 NVARCHAR(MAX) = CAST('' as nvarchar(max));
select @Tabela2 = cast((
    SELECT 
        td = ISNULL(CONVERT(VARCHAR(10), [Data de pagamento], 105), N''), '',
        td = ISNULL([Estado Pagamento], N''), '',
        td = ISNULL([Tipo Fornecimento], N''), '',
        td = ISNULL(Empresa, N''), '',
        td = ISNULL([Nome Entidade], N''), '',
        td = ISNULL(FORMAT(Montante, 'C', 'pt-PT'), N'0,00 €'), ''
    FROM dbo.Payments_Due_15Days
    ORDER BY [Data de pagamento] ASC
    FOR XML PATH('tr'), TYPE
) AS NVARCHAR(MAX));
```

And once again, the query result was added to the @OutputHTML variable in conjunction with the visual formatting:

```sql
    SET @OutputHTML = @OutputHTML + N'<h2>2. Pagamentos Previstos para a semana seguinte (8 a 15 Dias)</h2>    <table>        <tr>            <th>Data de Pagamento</th><th>Estado Pagamento</th><th>Tipo Fornecimento</th><th>Empresa</th><th>Nome Entidade</th><th>Montante</th>        </tr>' + ISNULL(@Tabela2, N'<tr><td colspan="6" style="color: #64748b; font-style: italic;">Nenhum registo encontrado.</td></tr>') + N'</table>';

```

The next table is the summary payments table for Company ID = 1. This table presented what turn out to be a challenge. One column has descriptions for each line and the other column has the aggregated value for that description, selected from the view [Resumo Pagamentos Company ID = 1]. To add each description and value to each respective column, the values of each line are cross applied to the columns . The conversion to xml is as in the previous columns. The addition to the @OutputHTML variable is also the same. 

```sql
  DECLARE @Tabela4 NVARCHAR(MAX) = CAST('' as nvarchar(max));
      SELECT @Tabela4 = cast((
      select
          td = PivotData.Campo, '',
          td = PivotData.Montante, ''
      FROM dbo.[Resumo Pagamentos Company ID =1] AS CC
      CROSS APPLY (VALUES         
          (N'Nº de pagamentos efectuados', CAST(CC.[Numero de pagamentos] AS NVARCHAR(50))),
          (N'Montante total de pagamentos à data', ISNULL(FORMAT(CC.[Montante Total de pagamentos efectuados à data], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante total a pagar nos próximos 7 dias', ISNULL(FORMAT(CC.[Montante Total Próximos 7 dias], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante total a pagar para a 2º semana', ISNULL(FORMAT(CC.[Montante Total para a semana seguinte], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante total a pagar nos próximos 15 dias', ISNULL(FORMAT(CC.[Montante Total Próximos 15 dias], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante total de pagamentos em atraso', ISNULL(FORMAT(CC.[Montante total de pagamentos em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante de prest. bancárias ou leasing em atraso', ISNULL(FORMAT(CC.[Montante de prestações bancárias e leasing em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante total de cons. bás. em atraso', ISNULL(FORMAT(CC.[Montante de consumos básicos em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante de TSU em atraso', ISNULL(FORMAT(CC.[Montante de TSU em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante de IRS em atraso', ISNULL(FORMAT(CC.[Montante de IRS em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante de IUC em atraso', ISNULL(FORMAT(CC.[Montante de IUC em atraso], 'C', 'pt-PT'), N'0,00 €')),
          (N'Montante de pag. relativos a viaturas em atraso', ISNULL(FORMAT(CC.[Montante de pagamentos com viaturas em atraso], 'C', 'pt-PT'), N'0,00 €'))
      ) AS PivotData(Campo, Montante)
      FOR XML PATH('tr'), TYPE
  ) AS NVARCHAR(MAX));

    SET @OutputHTML = @OutputHTML + N'<h2>3. Visão Geral de Pagamentos - Company ID = 1</h2>    <table>        <tr><th style="width: 60%;">Campo</th><th>Montante</th></tr>' + ISNULL(@Tabela4, N'<tr><td colspan="2">Sem dados.</td></tr>') + N'</table>';

```

The same process is applied to the summary table for Company ID = 2;

```sql
DECLARE @Tabela5 NVARCHAR(MAX) = CAST('' as nvarchar(max));
select @Tabela5 = cast ((
    SELECT 
        td = PivotData.Campo, '',
        td = PivotData.Montante, ''
    FROM dbo.[Resumo Pagamentos Company ID = 2] AS IP
    CROSS APPLY (VALUES         
        (N'Nº de pagamentos efectuados', CAST(IP.[Numero de pagamentos] AS NVARCHAR(50))),
        (N'Montante total de pagamentos à data', ISNULL(FORMAT(IP.[Montante Total de pagamentos efectuados à data], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante total a pagar nos próximos 7 dias', ISNULL(FORMAT(IP.[Montante Total Próximos 7 dias], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante total a pagar para a 2º semana', ISNULL(FORMAT(IP.[Montante Total para a semana seguinte], 'C', 'pt-PT'), N'0,00 €')), -- Corrigido aqui
        (N'Montante total a pagar nos próximos 15 dias', ISNULL(FORMAT(IP.[Montante Total Próximos 15 dias], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante total de pagamentos em atraso', ISNULL(FORMAT(IP.[Montante total de pagamentos em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante total de prestações bancárias em atraso', ISNULL(FORMAT(IP.[Montante de prestações bancárias e leasing em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante total de consumos básicos em atraso', ISNULL(FORMAT(IP.[Montante de consumos básicos em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante de TSU em atraso', ISNULL(FORMAT(IP.[Montante de TSU em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante de IRS em atraso', ISNULL(FORMAT(IP.[Montante de IRS em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante de IUC em atraso', ISNULL(FORMAT(IP.[Montante de IUC em atraso], 'C', 'pt-PT'), N'0,00 €')),
        (N'Montante de pag. relativos a viaturas em atraso', ISNULL(FORMAT(IP.[Montante de pagamentos com viaturas em atraso], 'C', 'pt-PT'), N'0,00 €'))
    ) AS PivotData(Campo, Montante)
    FOR XML PATH('tr'), TYPE
) AS NVARCHAR(MAX));

SET @OutputHTML = @OutputHTML + N'<h2>4. Visão Geral de Pagamentos - Company ID = 2</h2>    <table>        <tr><th style="width: 60%;">Campo</th><th>Montante</th></tr>' + ISNULL(@Tabela5, N'<tr><td colspan="2">Sem dados.</td></tr>') + N'</table>';

```

Finally, a footer is appended:

```sql
    SET @OutputHTML = @OutputHTML + N'    <br><hr style="border: 0; border-top: 1px solid #e2e8f0;">    <p style="font-size: 11px; color: #94a3b8; text-align: center; font-weight: 500;">        CrystalCompass - Sistema de Gestão Financeira ' + CONVERT(VARCHAR(20), GETDATE(), 120) + N'    </p>    </body>    </html>';

END;
GO
```

The html is stored in the @OutputHTML variable, that will be processed in a subsequent SSIS scrip task  to send for final automated email distribution trough a sql server job. 
