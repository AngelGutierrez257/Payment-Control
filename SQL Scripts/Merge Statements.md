
## Merge Statements

This file contains, explained, the merge statements used to clean the data and populate the dimension and fact tables. To each merge statement corresponds a Execute SQL Task in SSIS. 

### Merge to Dim Fornecedores

Each merge statement begins with a cte that handles nulls, converts the data if needed and adds a hash value to provide multi-column change detection in a efficient way that also simplifies the merge statement. In the hash statement, and given the low probability that these values register any change, all  attributes were included in the hash. the syntax is convert(varbinary(32), hashvalue('SHA2_256',concat ... and then list the columns, separated by '|' to  include in the hash. The hashvalue function converts the concatenated string to a hash in the SHA2_256 algorithm. The value is converted to varbinary(32) and is in that format that comparisons are made.
```sql
with
  cte_staging_fornecedores AS (
    select
      [Cód] AS Codigo,
      trim([Nome]) AS Nome,
      trim([Localidade]) AS Localidade,
      isnull([Telefone], 999999999) AS Telefone,
      isnull(trim([Desig.Comercial]), 'Nao definido') AS [Desig.Comercial],
      trim(NºContrib) AS [Nº Contribuinte],
      ImportDate AS DataInicio,
      1 AS IsCurrent,
      '9999-12-31' AS DataFim,
      convert(
        varbinary(32),
        hashbytes(
          'SHA2_256',
          CONCAT(
            [Cód],
            '|',
            [Nome],
            '|',
            [Localidade],
            '|',
            [Telefone],
            '|',
            [Desig.Comercial],
            '|',
            [NºContrib]
          )
        ),
        2
      ) AS HashValue
    FROM
      [fornecedores].[Staging_Fornecedores]
  )
MERGE
  [dbo].[Fornecedores] AS target using cte_staging_fornecedores AS source ON target.Codigo = source.[Codigo]
  AND target.IsCurrent = 1
WHEN MATCHED
  AND target.HashValue <> source.HashValue THEN
UPDATE SET
  target.IsCurrent = 0,
  target.DataFim = DATEADD(Day, -1, source.DataInicio)
WHEN NOT MATCHED BY target THEN
INSERT(
    Codigo,
    Nome,
    Localidade,
    Telefone,
    [Desig.Comercial],
    [Nº Contribuinte],
    DataInicio,
    DataFim,
    IsCurrent,
    HashValue
  )
VALUES
  (
    source.Codigo,
    source.Nome,
    source.Localidade,
    source.Telefone,
    source.[Desig.Comercial],
    source.[Nº Contribuinte],
    source.DataInicio,
    source.DataFim,
    source.IsCurrent,
    source.HashValue
  );


```



The merge is quite straightforward, defining source and target tables, or where the values come and where they are inserted, and defining that IsCurrent = 1 and the key that relates the rows. If both are equal and the hashvalue is changed, the ISCurrent value is changed to 0 and the present date is inserted in the EndDate column (DataFim). When the key is not found, the new values are inserted in the respective columns.


### Mergo to Dim Bancos (Banks)

The conversion of the banks names presented an issue, as it was requiring to cast manually to avoid character mis-conversion or misplacement. To make the conversion uniform and avoid duplication of code, a simple scalar-valued function was used to standardized name cleaning. The function takes the bank's name, uniforms them to upper case, handles nulls and misplaces characters or values and returns the variable with the first letter in Upper case and the rest in lower case.

```sql
USE [Gestao Financeira]
GO

CREATE function [dbo].[fn_ClassificacaoBancos] (@BancoPagamento nvarchar(256))
returns nvarchar(256)
as
begin
	declare @bn nvarchar(255);

	set @bn = trim(upper(@BancoPagamento));
	
	if @bn is null set @bn = 'Nao definido';

	if @bn like '%[0-9]%' set @bn = 'Nao definido';

	return upper(left(@bn,1))+ lower(substring(@bn,2, len(@bn)));
end;
GO

```

The function is called to populate the [Nome Banco] (Banks Name). As in the previous merge, the cte performs null handling, attributes default values and adds an hash for change tracking. The aim here is to mantain basic information on banks, to allow financial analysis by bank across companies. The merge is similar as the previous one.

```sql
WITH
  cte_bancos AS (
    SELECT DISTINCT	
     [dbo].[fn_ClassificacaoBancos]([Banco Pagamento])  AS [Nome Banco],
      'Nao definido' AS Contacto,
      999999999 AS [Contacto Telefonico],
      GETDATE() AS [DataInicio],
      '9999-12-31' AS [DataFim],
      1 AS IsCurrent,
      convert(
        varbinary(32),
        hashbytes('SHA2_256', isnull([dbo].[fn_ClassificacaoBancos]([Banco Pagamento]), 'Nao definido'))) AS HashValue
    FROM
      [dbo].[MapaPagamentos_Stagging]
    WHERE
      [Banco Pagamento] NOT LIKE '%[0-9]%' OR [Banco Pagamento] IS NULL
  )
MERGE
  [dbo].[Bancos] AS target using cte_bancos AS source ON target.[Nome Banco] = source.[Nome Banco]
  AND target.[IsCurrent] = 1
WHEN MATCHED
  AND target.HashValue <> source.HashValue THEN
UPDATE SET
  target.IsCurrent = 0,
  target.DataFim = DATEADD(Day, -1, source.DataInicio)
WHEN NOT MATCHED BY target THEN
INSERT(
    [Nome Banco],
    [Contacto],
    [Contacto Telefonico],
    [DataInicio],
    [DataFim],
    [IsCurrent],
    [HashValue]
  )
VALUES
  (
    source.[Nome Banco],
    source.[Contacto],
    source.[Contacto Telefonico],
    source.[DataInicio],
    source.[DataFim],
    source.[IsCurrent],
    source.[HashValue]
  );

```


### Merge to Dim Empresa (Dim Company)

In this case, has it is a dim with just two rows, a function was not used and the original code was kept. All the rest is as the previous merges.

```sql

with cte_Empresas as (
select distinct 
	case
		when trim(upper([Empresa])) is null then 'Nao indicado'
		else trim(upper([Empresa]))
	end as [Empresa],
	case when Empresa like '%Crystal%' then 'vat number'
		 when Empresa like '%Instal%' then 	'vat number' else 'Não definido' end AS NIF,
	Getdate() as [DataInicio],
	'9999-12-31' as [DataFim],
	1 as [IsCurrent],
	convert(
        varbinary(32),
        hashbytes('SHA2_256', isnull(trim(upper([Empresa])), 'Nao indicado'))
      ) AS HashValue
	  from [dbo].[MapaPagamentos_Stagging]
	  ),
	  format_empresa_cte as (
	  select upper(left([Empresa],1)) + lower(substring([Empresa],2,len([Empresa]))) as [Empresa], NIF, [DataInicio],[DataFim],IsCurrent,HashValue
  from cte_Empresas
	  )

merge 
		[dbo].[Empresa] as target using format_empresa_cte as source
		on target.Empresa = source.Empresa and target.IsCurrent = 1
	  when matched 
		and target.[HashValue] <> source.[HashValue] then
	update set
	  target.[IsCurrent] = 0,
	  target.[DataFim] = dateadd(day,-1,source.[DataInicio])
	when not matched by target then
	insert (
			[Empresa],
			[NIF],
			[DataInicio],
			[DataFim],
			[IsCurrent],
			[HashValue]
			)
	values (
			source.[Empresa],
			source.[NIF],
			source.[DataInicio],
			source.[DataFim],
			source.[IsCurrent],
			source.[HashValue]
			);



```


### Merge Dim Tipo Fornecimento (Supply Type)

As with the Dim Bancos Merge, a function was developed to make the cleanup code for the Supply Type name reusable and consistent. In this case, it is a simple case statement that cleans each type explicitly. That is important because the name of a given category could change over time, so any "bulk" process could fail or require more time to fix. Here, maintainability is easier. 

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER function [dbo].[fn_ClassificacaoTipoFornecimento] (@TipoFornecimento nvarchar(256))
returns nvarchar(256)
as
begin
	declare @tf nvarchar(256) = trim(upper(@TipoFornecimento));

	return (
	case
        WHEN @tf LIKE '%SEGURO%' THEN 'Seguros'
        WHEN @tf LIKE '%VIDRO%' THEN 'Vidro'
		WHEN @tf LIKE '%INFOR%' THEN 'Informatica'
		WHEN @tf LIKE '%CONTABILIDADE%' THEN 'Contabilidade'
		WHEN @tf LIKE '%VIATURAS%' THEN 'Viaturas'
		WHEN @tf LIKE '%VEICULOS%' THEN 'Viaturas'
		WHEN @tf LIKE '%CONSUMOS BASICOS%' then 'Consumos Básicos'
		WHEN @tf LIKE '%LEASING%' or @tf LIKE '%EMPRÉSTIMO%' or @tf LIKE '%Emprestimo%' then 'Empréstimos Bancários & Leasing'
		WHEN @tf LIKE '%PPC%' then 'Pagamento por Conta'
		WHEN @tf LIKE '%REPARAÇÂO%' or @tf LIKE '%REPARACAO%' or @tf LIKE '%REPAR%' then 'Reparação'
		WHEN @tf LIKE '%SALARIOS%' or @tf LIKE '%SALÁRIOS%' then 'Salários'
		WHEN @tf LIKE '%MAQUINA%' then 'Maquinaria'
		WHEN @tf LIKE '%MATERIAL COLOCACAO%' or @tf like '%MATERIAL COLOCAÇÃO%' or @tf like '%MATERIAL COLOCACÃO%' then 'Material Colocação'
		WHEN @tf LIKE '%TELECOM%' THEN 'Telecomunicações'
		WHEN @tf LIKE '%S.S.%' THEN 'Segurança Social'
		WHEN @tf LIKE '%ElECTRICIDADE%' or @tf LIKE '%ELETRICIDADE%' then 'Eletricidade'
        ELSE isnull(@tf, 'Nao definido')
	end
	);
end;
GO


```

After that, the cte and merge are similar as the previous ones.

```sql
WITH
  cte_TipoFornecimento AS (
    SELECT DISTINCT
    [dbo].[fn_ClassificacaoTipoFornecimento]([Tipo Fornecimento]) as [Tipo Fornecimento Limpo],
      getdate() AS [Data de Inicio],
      '9999-12-31' AS [Data de Fim],
      1 AS [IsCurrent],
      convert(
        varbinary(32),
        hashbytes(
          'SHA2_256',
           [dbo].[fn_ClassificacaoTipoFornecimento]([Tipo Fornecimento]))) AS [hashbytes]
    FROM
      [dbo].[MapaPagamentos_Stagging]
  )
MERGE
  [dbo].[Tipo Fornecimento] AS target using (select * from ( select *, row_number() over (partition by [Tipo Fornecimento Limpo] Order by  [Data de Inicio] desc ) as RowID from cte_TipoFornecimento) t where RowID = 1) AS source ON target.[Tipo Fornecimento] = source.[Tipo Fornecimento Limpo]
  AND target.[IsCurrent] = 1
WHEN MATCHED
  AND target.[hashbytes] <> source.[hashbytes] THEN
UPDATE SET
  target.[IsCurrent] = 0,
  target.[Data de Fim] = DATEADD(Day, -1, source.[Data de Inicio])
WHEN NOT MATCHED BY target THEN
INSERT(
    [Tipo Fornecimento],
    [Data de Inicio],
    [Data de Fim],
    [IsCurrent],
    [hashbytes]
  )
VALUES
  (
    source.[Tipo Fornecimento Limpo],
    source.[Data de Inicio],
    source.[Data de Fim],
    source.[IsCurrent],
    source.[hashbytes]
  );


```


### Merge Dim Tipo Entidade (Entity Type)

In this merge, 2 functions were used. As both the type name and the type description require uniformization and are used in more than one place, functions were used to clean both.  The first function, [fn_ClassificacaoTipoEntidade], is similar to the previous one, with a case statement to clean each type name. The case was used for the same reasons. 

```sql

USE [Gestao Financeira]
GO

CREATE OR ALTER function [dbo].[fn_ClassificacaoTipoEntidade] (@TipoEntidade nvarchar(256))
returns nvarchar(256)
as
begin
	declare @te nvarchar(255) = trim(upper(@TipoEntidade));

	return (
	case
		when @te like '%COMBÚSTIVEL%' or @te like '%COMBUSTÍVEL%' or @te like '%COMBUS%' then 'Combustivel'
		when @te like '%BANCO%' then 'Banco'
		when @te like '%ESTADO%' then 'Estado'
		when @te like '%FORNECEDOR MATERIAL%' then 'Fornecedor Material'
		when @te like '%FORNECEDOR MAQ%' then 'Fornecedor Maquinaria e Peças'
		when @te like '%RECURSO%' then 'Recursos Humanos'
		when @te like '%SERVIÇOS%' or @te like '%SERVICOS%' then 'Serviços'
		when @te like '%VIDRA%' then 'Vidraria'
		when @te like '%ENERG%' then 'Energia'
		when @te like '%SEGUR%' then 'Seguros'
		when @te like '%NÂO DEFINIDO%' or @te like '%NAO DEF%' then 'Não definido'
		when @te like '%VIATU%' then 'Viaturas'
		when @te like '%TESOUR%' then 'Tesouraria'
		when @te like '%S.S.%' or @te like '%SEG. SOCIAL%' or @te like '%SEGURANCA SOCIAL%' or @te like '%SEGURANÇA SOCIAL%' then 'Segurança Social'
		else isnull(@te, 'Valor ausente')
	end
	);
end;
GO

```

The second function, [fn_DescricaoTipoEntidade], adds a description to each category name. Again, the main advantage of functions in this case is it reusability, as this is a verbose case statement.

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER   function [dbo].[fn_DescricaoTipoEntidade] (@TipoEntidade nvarchar(500))
returns nvarchar(500)
as
begin
	declare @te nvarchar(500) = trim(upper(@TipoEntidade));

	return (
	case
		when @te like '%COMBÚSTIVEL%' or @te like '%COMBUSTÍVEL%' or @te like '%COMBUS%' then 'Compra de combustíveis, óleos e produtos associados'
		when @te like '%BANCO%' then 'Gastos com juros, comissões e empréstimos bancários'
		when @te like '%ESTADO%' then 'Pagamento de impostos, taxas e multas ao Estado'
		when @te like '%FORNECEDOR MATERIAL%' then 'Fornecedor de peças e ferragens para colocação de vidro e material associado'
		when @te like '%FORNECEDOR MAQ%' then 'Compra deconsumíveis, acessórios e material usado para transformar vidro'
		when @te like '%RECURSO%' then 'Salários, empréstimos a trabalhadores e outros encargos com funcionários'
		when @te like '%SERVIÇOS%' or @te like '%SERVICOS%' then 'Serviços de contabilidade, consultoria, comunicações, legais e outros.'
		when @te like '%VIDRA%' then 'Fornecimento de vidro em chapa ou vidro transformado e derivados para colocação ou comercialização'
		when @te like '%ENERG%' then 'Provedores de energia (eletricidade e similares, exclui combustível (diesel, gasolina e similares)).'
		when @te like '%SEGUR%' then 'Provedores de seguros auto, imobiliário, acidentes de trabalho e similares.'
		when @te like '%NÂO DEFINIDO%' or @te like '%NAO DEF%' then 'Tipo de entidade não especificado na origem'
		when @te like '%VIATU%' then 'Engloba as despesas de manutenção, reparação e uso de veículos que não sejam combustível, seguros e financiamento.'
		when @te like '%TESOUR%' then 'Serviços financeiros de gestão de caixa como contas caucionadas, confirming, letras, etc.'
		when @te like '%S.S.%' or @te like '%SEG. SOCIAL%' or @te like '%SEGURANCA SOCIAL%' or @te like '%SEGURANÇA SOCIAL%' then 'Obrigações com Seg. Social, TSU e prestações similares'
		else isnull(@te, 'Valor não especificado')
	end
	);
end;
GO

```

Finally, the two functions are used to populate the columns [Tipo Entidade] (entity type) and [Descricao] (description), and the rest of the code is similar as the other merge statements. 


```sql
with cte_TipoEntidade as (
select distinct 
[dbo].[fn_ClassificacaoTipoEntidade]([Tipo Entidade]) as [Tipo Entidade],
	[dbo].[fn_DescricaoTipoEntidade]([Tipo Entidade]) as [Descricao],
				getdate() as [DataInicio],
				'9999-12-31' as [DataFim],
				1 as [IsCurrent],
				convert(varbinary(32), hashbytes('SHA2_256', 
				[dbo].[fn_ClassificacaoTipoEntidade]([Tipo Entidade]) 				)) as [HashValue]
				from [dbo].[MapaPagamentos_Stagging]
	)

merge [dbo].[Tipo Entidade] as target using cte_TipoEntidade as source
	on target.[Tipo Entidade] =  source.[Tipo Entidade] 
	and target.IsCurrent = 1
	when matched and target.[HashValue] <> source.[HashValue] then
	update set
	target.IsCurrent = 0,
	target.[DataFim] = dateadd(day,-1,source.[DataInicio])

	when not matched by target then
	insert (
		[Tipo Entidade],
		[Descricao],
		[DataInicio],
		[DataFim],
		[IsCurrent],
		[HashValue]
		)
	values (
		source.[Tipo Entidade],
		source.[Descricao], 
		source.[DataInicio], 
		source.[DataFim], 
		source.[IsCurrent], 
		source.[HashValue]
		);
```


### Merge Tipo Pagamento (Payment type)

This merge is very similar to the merge Tipo Entidade. Uses two functions, as well. [fn_ClassificacaoTipoPagamento] cleans the various payment types names, similarly with a case statement.  In this case, declaring the variable to @fp nvarchar(255) = trim(upper(@FormaPagamento)) would be preferable to avoid repeating it at any line, but were left there and the code will be updated.

```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER function [dbo].[fn_ClassificacaoTipoPagamento] (@FormaPagamento nvarchar(255))
returns nvarchar(256)
as
begin
	--declare @fp nvarchar(255) = trim(upper(@FormaPagamento));

	return (
	case 
	when trim(upper(@FormaPagamento)) like '%Ref%' then 'Entidade Referencia'
	when (trim(upper(@FormaPagamento)) like '%Debito%' or trim(upper(@FormaPagamento)) like '%Débito%') then 'Debito Directo'
	when trim(upper(@FormaPagamento))  collate Latin1_General_CI_AI like '%Transferencia%'  then 'Transferencia SEPA' -- collate Latin1_General_CI_AI makes Transferencia case and accent insensitive
	when trim(upper(@FormaPagamento)) like '%Estado%' then 'Estado'
	when trim(upper(@FormaPagamento)) like '%NUMERARIO%' then 'Numerário'
	when trim(upper(@FormaPagamento)) like '%NAO DEFINIDO%' then 'Não definido'
	when trim(upper(@FormaPagamento)) like '%Serviços%' then 'Serviços'
	else isnull(trim(upper(@FormaPagamento)), 'Valor ausente') 
	end 
	);
end;
GO

```

The second function, [fn_DescricaoTipoPagamento],  performs the same transformation as the function [fn_DescricaoTipoEntidade]. Adds a comment to each payment type. 
```sql
USE [Gestao Financeira]
GO

CREATE OR ALTER   function [dbo].[fn_DescricaoTipoPagamento] (@FormaPagamento nvarchar(500))
returns nvarchar(500)
as
begin
	--declare @fp nvarchar(500) = trim(upper(@FormaPagamento));

	return (
	case 
	when trim(upper(@FormaPagamento)) like '%Ref%' then 'Pagamento efectuado por entidade e referencia'
	when (trim(upper(@FormaPagamento)) like '%Debito%' or trim(upper(@FormaPagamento)) like '%Débito%') then 'Pagamento efectuado por débito directo'
	when trim(upper(@FormaPagamento))  collate Latin1_General_CI_AI like '%Transferencia%'  then 'Pagamento efectuado por transferencia bancaria'  -- collate Latin1_General_CI_AI makes Transferencia case and accent insensitive
	when trim(upper(@FormaPagamento)) like '%Estado%' then 'Pagamento efectuado ao Estado'
	when trim(upper(@FormaPagamento)) like '%NUMERARIO%' then 'Pagamento efectuado em numerário.'
	when trim(upper(@FormaPagamento)) like '%NAO DEFINIDO%' then 'Forma de pagamento não determinada ou posposta.'
	when trim(upper(@FormaPagamento)) like '%Serviços%' then 'Pagamento efectuado através de plataforma prestadores de serviços'
	else isnull(trim(upper(@FormaPagamento)),'Forma de pagamento não inserida' ) 
	end 
	);
end;
GO
```

As with the previous merges, the cte and merge function the same way. The functions are called to populate the type name and description columns int he cte, the same transformations are performed and the merge works as the previous ones.


```sql
with cte_TipoPagamento as (
select distinct 
[dbo].[fn_ClassificacaoTipoPagamento]([Forma Pagamento]) as [Tipo de Pagamento],
[dbo].[fn_DescricaoTipoPagamento]([Forma Pagamento]) as Descricao,
	GETDATE() as DataInicio,
	'9999-12-31' as DataFim,
	1 as IsCurrent,
	convert(varbinary(32), hashbytes('SHA2_256', 
[dbo].[fn_ClassificacaoTipoPagamento]([Forma Pagamento]) 		)) as [HashValue]
		from [dbo].[MapaPagamentos_Stagging]
)

Merge [dbo].[Tipo Pagamento] as target using cte_TipoPagamento as source 
	on target.[Tipo de Pagamento] = source.[Tipo de Pagamento]
	and target.IsCurrent = 1

when matched
	and target.HashValue <> source.HashValue then
	update set
	target.[IsCurrent] = 0,
	target.[DataFim] = DATEADD(Day, -1, source.[DataInicio])
when not matched by target then
	insert( 
		[Tipo de Pagamento],
		[Descricao],
		[DataInicio],
		[DataFim],
		[IsCurrent],
		[HashValue])
values (
		source.[Tipo de Pagamento],
		source.[Descricao],
		source.[DataInicio],
		source.[DataFim],
		source.[IsCurrent],
		source.[HashValue]
		);
```


### Merge Fact Pagamentos (Fact Payments)

This is the final merge. The first cte re-uses the previous functions to populate the respective columns. They are also used to define the hash for each fact row. In this case, as the simpler process of the other merges was not providing the correct results, the hash was divided between a business hash and a row hash. The business hash relates only to the business values. The control row hash represents the current state of the row, having the attributes that characterize the row at a given time.  

The second cte, cte_pagamentos, joins the fact table to the dim tables and uses the values provided by the functions to confirm its correct. The final cte ensures that only one row per payment exists. The merge updates when new information is added to existing rows and, if the keys are not found, inserts new rows.

```sql
with cte_stagging as (
 select [ImportID] , 
 [Data], 
 dbo.fn_ClassificacaoBancos([Banco Pagamento]) as NomeBanco, -- use of function to standradized dim merge logic with fact merge logic. Increases performance and avoids logic drifting
 isnull(trim(upper([Nome Entidade])),'Nao definido') as [NomeEntidade],
 dbo.fn_ClassificacaoTipoEntidade([Tipo Entidade]) as TipoEntidade, -- use of function to standradized dim merge logic with fact merge logic. Increases performance and avoids logic drifting
 dbo.fn_ClassificacaoTipoFornecimento([Tipo Fornecimento]) as TipoFornecimento, -- use of function to standradized dim merge logic with fact merge logic. Increases performance and avoids logic drifting
 dbo.fn_ClassificacaoTipoPagamento([Forma Pagamento]) as TipoPagamento, -- use of function to standradized dim merge logic with fact merge logic. Increases performance and avoids logic drifting
 isnull(trim(upper([Empresa])),'Nao definido') as Empresa,
 [Descricao],
 [Banco de debito Previsto],
 [Valor],
 [Em atraso?] as [Estado Pagamento],
 [Data Situação actual], 
 getdate() as LoadDate, 
 convert(varbinary(32), hashbytes('SHA2_256', concat([Data], '|',isnull(trim(upper([Nome Entidade])),'Nao definido'),'|',
 dbo.fn_ClassificacaoTipoEntidade([Tipo Entidade]), '|',dbo.fn_ClassificacaoTipoFornecimento([Tipo Fornecimento]),'|',
 dbo.fn_ClassificacaoTipoPagamento([Forma Pagamento]), '|',dbo.fn_ClassificacaoBancos([Banco Pagamento]), '|',
 isnull(trim(upper([Empresa])),'Nao definido'), '|',[Valor]))) as [BusinessHashValue],
 convert(varbinary(32), hashbytes('SHA2_256', concat([Data], '|',isnull(trim(upper([Nome Entidade])),'Nao definido'),'|',
 dbo.fn_ClassificacaoTipoEntidade([Tipo Entidade]), '|',dbo.fn_ClassificacaoTipoFornecimento([Tipo Fornecimento]),'|',
 dbo.fn_ClassificacaoTipoPagamento([Forma Pagamento]), '|',dbo.fn_ClassificacaoBancos([Banco Pagamento]), '|',
 isnull(trim(upper([Empresa])),'Nao definido'), '|',[Valor], '|', isnull(trim(upper([Em atraso?])),'Nao definido'), '|',[Data Situação actual])))  
 as [ControlChangeRowHash] -- controls changes to payment state and compares to the same payment in the busines hash value
 from [dbo].[MapaPagamentos_Stagging] where [Data] is not null
 ),
 cte_pagamentos as (
		select 
		a.[ImportID],
		a.[Data],
		isnull(b.[ID],6) as BancoID,
		isnull(c.[ID],3) as EmpresaID,
		isnull(d.[ID],1002) as FornecedorID,
		isnull(e.ID,5)  as [Tipo Entidade ID],
		isnull(f.ID,10) as[Tipo Fornecimento ID],
		isnull(h.ID,4) as [Tipo Pagamento ID],
		a.[NomeEntidade],
		a.[Descricao],
		a.[Valor],
		a.[Banco de debito Previsto],
		a.[Estado Pagamento],
		a.[Data Situação actual],
		a.[LoadDate],
		a.[BusinessHashValue],
		a.[ControlChangeRowHash]
		from cte_stagging a
		left join [dbo].[Bancos] b on b.[Nome Banco] = a.NomeBanco and b.IsCurrent = 1
		left join [dbo].[Empresa] c on c.[Empresa] = a.Empresa and c.[IsCurrent] = 1
		left join [dbo].[Fornecedores] d on trim(upper(d.[Nome])) = trim(upper(a.NomeEntidade)) and d.[IsCurrent] = 1
		left join [dbo].[Tipo Entidade] e on e.[Tipo Entidade] = a.TipoEntidade and e.[IsCurrent] = 1
		left join [dbo].[Tipo Fornecimento] f on f.[Tipo Fornecimento] = a.TipoFornecimento and f.[IsCurrent] = 1
		left join [dbo].[Tipo Pagamento] h on h.[Tipo de Pagamento] = a.TipoPagamento and h.[IsCurrent] = 1
		),
		cte_pagamentos_dedup as (
		select * from (select *, row_number() over (partition by [BusinessHashValue] order by [ImportID] desc ) as rownumber from cte_pagamentos ) as [rowcount] where rownumber =1
		)
		--select * from cte_pagamentos_dedup

	merge [dbo].[Pagamentos] as target using cte_pagamentos_dedup as source 
		on target.[BusinessHashValue] = source.[BusinessHashValue]
	when matched and target.[ControlChangeRowHash] <> source.[ControlChangeRowHash] then
	update set
		target.Data = source.Data,
		target.BancoID = source.BancoID,
		target.[FornecedorID] = source.FornecedorID,
		target.[Tipo Fornecimento ID] =source.[Tipo Fornecimento ID],
		target.[Tipo Pagamento ID] = source.[Tipo Pagamento ID],
		target.[Tipo Entidade ID] = source.[Tipo Entidade ID],
		target.[EmpresaID] = source.EmpresaID,
		target.[Nome Entidade] =  source.[NomeEntidade],
		target.[Descricao] = source.[Descricao],
		target.[Valor] = source.[Valor],
		target.[Banco de debito Previsto] = source.[Banco de debito Previsto],
		target.[Estado Pagamento] = source.[Estado Pagamento],
		target.[Data Situação actual]  =source.[Data Situação actual],
		target.[LoadDate] =  source.[LoadDate],
		target.[ControlChangeRowHash] = source.[ControlChangeRowHash]
	when not matched by target then
	insert (
		[ImportID],
		[Data],
		[BancoID],
		[FornecedorID],
		[Tipo Fornecimento ID],
		[Tipo Pagamento ID],
		[Tipo Entidade ID],
		[EmpresaID],
		[Nome Entidade],
		[Descricao],
		[Valor],
		[Banco de debito Previsto],
		[Estado Pagamento],
		[Data Situação actual],
		[LoadDate],
		[BusinessHashValue],
		[ControlChangeRowHash]
		)
		values (
		source.[ImportID],
		source.[Data],
		source.[BancoID],
		source.[FornecedorID],
		source.[Tipo Fornecimento ID],
		source.[Tipo Pagamento ID],
		source.[Tipo Entidade ID],
		source.[EmpresaID],
		source.[NomeEntidade],
		source.[Descricao],
		source.[Valor],
		source.[Banco de debito Previsto],
		source.[Estado Pagamento],
		source.[Data Situação actual],
		source.[LoadDate],
		source.[BusinessHashValue],
		source.[ControlChangeRowHash]
		);
		

```
