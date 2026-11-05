
## Merge Statements

This file contains, explained, the merge statements used to clean the data and populate the dimension and fact tables. To each merge statement corresponds a Execute SQL Task in SSIS. 

### Merge to Dim Fornecedores

Each merge statement begins with a cte that handles nulls, converts the data if needed and adds a hash value to provide multi-column change detection in a efficient way and simplifies the merge statement. In the hash statement, and given the low probability that these values register any change, all  attributes were included in the hash. the syntax is convert(varbinary(32), hashvalue('SHA2_256',concat ... and then list the columns, separated by '|' to  include in the hash. The hashvalue function converts the concatenated string to a hash in the SHA2_256 algorithm. The value is converted to varbinary(32) and is in that format that comparisons are made.
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

