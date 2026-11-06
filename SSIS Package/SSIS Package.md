
## SSIS Package Payment Control

This file explains step by step how the ssis package works and what is configuration is. 

The package extracts data from two excel sources. MapaPagmentos, where the MapaPagamentos2026 sheets is extracted. Its columns and data types are:

| column name | column data type | column description |
| :--- | :--- | :--- |
| Ano | Number | The year of the date of payment |
| Mes | Number | The month of the date of payment |
| Dia | Date | The day of the date of payment |
| Dia da semana | text | The name of the day of the date of payment |
| Codigo Fornecedor | Number |  Code for the supplier |
| Nome Entidade | General | Name of the Supplier |
| Tipo Entidade | general |  Type of the entity |
| Tipo Fornecimento | general | Category of supply |
| Empresa | general |  business unit |
| Descrição | general | Description  of the transaction |
| Foram de Pagamento | general | Method of payment used |
| Valor | number | Monetary amount  |
| Banco de Debito | general | Bank account used for the debit transaction |
| Obs Adicionais | general | Additional observations or remarks |
| Entidade beneficiada confirming | general | Details regarding the confirming beneficiary entity - for payments using confirming only |
| Em atraso? | general | Status indicating whether the payment is overdue |
| Situacao actual | general | Current status of the record |
| Data Situação actual | date | Date of the current status update |
| Banco Pagamento | General | Bank used to execute the payment |

The second excel file, Lista Fornecedores (suppliers list), is extracted directly from the erp program. It contains basic information on suppliers:

| column name | column data type | column description |
| :--- | :--- | :--- |
| Cód. | general | Code for the entity |
| Nome | General | Full name of the supplier |
| Localidade | general | Supplier's address |
| Telefone | general | Contact telephone number |
| Desig.Comercial | general | Commercial or business name of the supplier, if applicable |
| NºContribuinte | general | Tax identification number |
| E | general | Not used |
| A | general | Not used |

