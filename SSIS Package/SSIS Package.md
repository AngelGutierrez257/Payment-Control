
# SSIS Package Payment Control

This file explains step by step how the ssis package works and what is configuration is. 

## Sources
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


## Variables

The variables defined for the package are the following, each one will be explained in the context of the tasks where they are used.

| Variable Name | Variable Scope | Data Type | Value |
| :--- | :--- | :--- | :--- |
| appPass | Package | String | |
| Destination_Excel_MapaPagamentos | Package | String | D:\<file path>|
| HtmlReportBody | Package | String | |
| LoadDate | Package | DateTime | |
| Path_MapaPagamentos_Copia | Package | String | |
| RunID | Package | String | |
| Source_Excel_MapaPagamentos | Package | String | D:\<file path> |

<img width="1915" height="289" alt="image" src="https://github.com/user-attachments/assets/9ecadb03-e8d1-4dcd-9409-21e12bc858aa" />

## Parameters
The two main parameters used are appPass and SendingEmail. For security reasons, the appPass string is not present. Is is the password for the email service that you use to send the email and is used in the Send Email task. Sending Email is just the email address for that email. You can also parameterize the receiving emails. You should, in fact. To work, the email package parameters need to be sensitive = false in the local machine.

![[Pasted image 20261012182032.png]]


## Connectors

The package requires 3 connectors, very straightforward. The first is an excel connector to the MapaPagamentos file. Everything was kept as it is, except the selection of the sheet, that goes to the specific sheet needed. The connector is set to standard MS Excel 2007-201. Another excel connector is defined with the path to the Lista de Fornecedores file. No sheets selection is needed as it has just 1 sheet. Finally, a OLE DB connector is defined to a sql server instance installed in the local machine. It uses windows authentication and connects to the Gestao Financeira database.

## Tasks

### Task 1. LoadDate Expression Task

The first task does the loading of the LoadDate variable, which is, as the picture shows, set to GETDATE() (the date when the run occurs).

<img width="745" height="747" alt="image" src="https://github.com/user-attachments/assets/63710473-5694-42cd-9e4a-2247b23ea27f" />

### Task 2. Copy MapaPagamentos Scrip Task.

This task aims to solve one of the issues of excel files, which is the lock Windows makes in open office documents. In production environments sometimes the file stays open and the package does not run. So, to avoid that, a copy of is made using a script task. 

First, the script requires a source, destination and date variables, defined above: User::Destination_Excel_MapaPagamentos,User::LoadDate,User::Source_Excel_MapaPagamentos 
You can select them in the drop down menu that appears after clicking the ... in the ReadOlnyVariables field.

<img width="843" height="771" alt="Copy MapaPagamentos Script Task" src="https://github.com/user-attachments/assets/f10464b2-0c7d-49c9-bd29-85223140276d" /> 


Click in edit script and there you develop the code in C#. The code can be divided in two parts:

In the first one, public partial class ScriptMain, we define the code that will execute in the Main(). In this case, they define the access to the excel file as read and the sharing operations allowed to continue for the shared file. They do not bypass the lock, has it cannot alter the file. It is not needed, as it is needed to copy the contents of the sheet.

```C#

using Microsoft.Win32.SafeHandles;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;

namespace ST_ScriptTask
{
    [Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        }

        private bool _fireAgain = true;

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern SafeFileHandle CreateFile(
            string lpFileName, uint dwDesiredAccess, uint dwShareMode,
            IntPtr lpSecurityAttributes, uint dwCreationDisposition,
            uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        const uint GENERIC_READ = 0x80000000;
        const uint FILE_SHARE_READ = 0x1;
        const uint FILE_SHARE_WRITE = 0x2;
        const uint FILE_SHARE_DELETE = 0x4;
        const uint OPEN_EXISTING = 3;


```

In the Main(), the variables are internalized in local variables, the directory is created if it is not found, the file specified in the sourcePath is read with GENERIC_READ, the operations that are allowed are also expressed, OPEN_EXISTING keyword defines that the file must exist or otherwise throw the error defined in the if clause. The copy in itseld is done with using (FileStream fsDestination = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                    {
                        fsSource.CopyTo(fsDestination);
                    }
    The rest of the code are standard Dts methods for execution and error messages.

  ```C#
    public void Main()
        {
            try
            {
                string sourcePath = (string)Dts.Variables["User::Source_Excel_MapaPagamentos"].Value;
                string destinationPath = (string)Dts.Variables["User::Destination_Excel_MapaPagamentos"].Value;
                DateTime loadDate = (DateTime)Dts.Variables["User::LoadDate"].Value;

                string destFileName = "MapaPagamentos2026.xlsx";
                string destPath = Path.Combine(destinationPath, destFileName);

                if (!Directory.Exists(destinationPath))
                {
                    Directory.CreateDirectory(destinationPath);
                }

                SafeFileHandle handle = CreateFile(
                    sourcePath,
                    GENERIC_READ,
                    FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero);

                if (handle.IsInvalid)
                {
                    int errorCode = Marshal.GetLastWin32Error();
                    throw new IOException($"Cannot open source file '{sourcePath}'. Win32 error code: {errorCode}");
                }

                using (FileStream fsSource = new FileStream(handle, FileAccess.Read))
                {
                    if (fsSource.Length == 0)
                    {
                        throw new IOException($"Source file '{sourcePath}' is 0 bytes. Aborting copy.");
                    }

                    using (FileStream fsDestination = new FileStream(destPath, FileMode.Create, FileAccess.Write))
                    {
                        fsSource.CopyTo(fsDestination);
                    }
                }

                Dts.Events.FireInformation(
                    0,
                    "Copy MapaPagamentos2026",
                    $"Copied '{sourcePath}' to '{destPath}' using LoadDate {loadDate}.",
                    string.Empty,
                    0,
                    ref _fireAgain);

                Dts.TaskResult = (int)ScriptResults.Success;
            }
            catch (Exception ex)
            {
                Dts.Events.FireError(
                    0,
                    "Copy Excel File",
                    ex.Message,
                    string.Empty,
                    0);

                Dts.TaskResult = (int)ScriptResults.Failure;
            }
        }
    }
}

  ```

### Import Raw Data Container

<img width="823" height="291" alt="Import Raw Data Container" src="https://github.com/user-attachments/assets/6087c5f7-36af-4b3f-9749-7d827a25e805" />


The first container, of type sequence, holds the tasks that import the files. Before the data flow tasks that import both excel sheets, the import tables are truncated with simple truncate table statements. All the rest is kept as default.

<img width="843" height="774" alt="truncate fornecedores" src="https://github.com/user-attachments/assets/b13c36a8-2045-49a4-94e4-784d799f0ba7" />

<img width="845" height="771" alt="Truncate MapaPagamentos" src="https://github.com/user-attachments/assets/a6afa724-5137-4979-ab86-34fb12f59119" />

The Truncate Fornecedores sql task connects to the Import Excel Data Fornecedores Stagging: 

<img width="590" height="358" alt="Import Excel to Fornecedores Stagging Data Flow" src="https://github.com/user-attachments/assets/44be694b-a404-4142-9046-46d4ee2e265a" />

This Data Flow task, very common and simple, is composed of an Excel Source, a derived column transformation and an OLE DB Destination. 

The excel source is defined and mapped as follows:

<img width="846" height="837" alt="ListaFornecedoresExcelSource_Connector" src="https://github.com/user-attachments/assets/67137ad2-d295-43b1-854a-bf2af1ff4960" />

<img width="847" height="838" alt="ListaFornecedoresExcelSource_Mapping" src="https://github.com/user-attachments/assets/62eee41d-10de-4e7f-8f6c-f034cb5b3be2" />

The derived column transformations adds a new column, ImportDate, to the Lista Fornecedores table before is imported to the db, using the LoadDate variable:

<img width="885" height="832" alt="ListaFornecedores_DerivedColumn_AddDate" src="https://github.com/user-attachments/assets/49c21094-24ca-4644-add8-9ed7480ed76f" />

Finally, is uploaded to the import table and the columns mapped to the destination, which create scripts are in the SQL Scripts folder:

<img width="846" height="837" alt="ListaFornecedoresExcelSource_Connector" src="https://github.com/user-attachments/assets/9723c88d-601c-48e8-8956-6a94175412a7" />

<img width="847" height="838" alt="ListaFornecedoresExcelSource_Mapping" src="https://github.com/user-attachments/assets/e60fa94c-abf5-4157-b7b9-860751c5ebd1" />










