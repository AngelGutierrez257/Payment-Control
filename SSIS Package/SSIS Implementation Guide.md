
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

### Import Raw Data Sequence Container

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

The Data Flow Task that imports the MapaPagamentos sheet to the Gestao Financeira DB is similar to the previous one, but refleting the more complex strucutre of the imported sheet: 

<img width="708" height="344" alt="Import Excel to MapaPagamentos Stagging Data Flow" src="https://github.com/user-attachments/assets/a82e3756-7d4a-419d-ba27-a027af9612d3" /> 

The Excel Source transformations is configured as follows: 

<img width="849" height="836" alt="MapaPagamentos_ExcelSource_ConnectorPage" src="https://github.com/user-attachments/assets/12e9fcbf-3c3d-4d10-9f47-d09d222c1f4c" />

<img width="848" height="842" alt="MapaPagamentos_ExcelSource_Mapping" src="https://github.com/user-attachments/assets/0d7e40c1-35d7-459e-bb2f-33ad3569c32c" />

The derived column transformation works similarly to the previous derived column, adding the column ImportDate using the LoadDate variable:

<img width="886" height="835" alt="MapaPagamentos LoadDate Derived Column" src="https://github.com/user-attachments/assets/81e48b92-14e9-475b-a4ac-0be941e8e5c9" /> 

Then it is imported with an OLE DB Destination transformation, only declaring the connector, the destination table and the column mapping:

<img width="850" height="835" alt="MapaPagamentos OleDB Destination " src="https://github.com/user-attachments/assets/f8c58e6e-4c39-48b9-9e4a-ceae2f0d95c3" />

<img width="841" height="838" alt="MapaPagamentos OleDB Destination Mappings " src="https://github.com/user-attachments/assets/bdb369d2-ade8-4613-ac2d-8c12cd0ba381" />


### Merge Statements Sequence Container

The next container, that contains the merge statements executed in a specific order, starting with merging to the Suppliers Dimension (Lista Fornecedores) and consequently from the simplest merge to the merge for the fact table. The SQL Scripts are shown and explained in the SQL Scripts folder, Merge Statements.md file. Here the container and the execute sql tasks contained in it will be briefly explained. One point is that, as best practice, the execution of merges should be carried out by executing store procedures and not putting the complete merge statement in the siss task, but the urgency in put this pipeline into production led me to focus more on guaranteeing that the data given is reliable and trustworthy, and this aspect was overlooked. 

<img width="820" height="431" alt="Merge Statements Container 1" src="https://github.com/user-attachments/assets/2a1873b4-8c84-4ab0-9ce4-69943b690254" /> 
<img width="825" height="324" alt="Merge Statements Container 2" src="https://github.com/user-attachments/assets/9711388e-580a-4227-99fa-f2bca6d1f722" />

The first Execute SQL Task is the Execute merge to Dim Fornecedores. All this tasks work the same. They use the OLE DB Connector to execute the sql statement contained in the SQL Statement field:

<img width="843" height="773" alt="Execute Merge to Dim Fornecedores" src="https://github.com/user-attachments/assets/ccab258d-69fd-47df-9940-d2bec22155bd" /> 

The second merge related task is the merge of Dim Bancos:

<img width="843" height="770" alt="Execute Merge to Dim Bancos" src="https://github.com/user-attachments/assets/cee754b5-6707-4f10-97f4-a35c8c4d0636" />

 The third merge related task is the merge to Dim Empresas. Has with the previous merge tasks, the only actions were defining the connector and adding the sql statement to execute:

 <img width="844" height="737" alt="Execute Merge to Dim Empresa" src="https://github.com/user-attachments/assets/d89e7b67-d18a-4412-ac8e-9a6206d3da02" />

 The fourth merge statement is the merge to Dim Tipo Fornecimento: 

 <img width="843" height="773" alt="Execute Merge to Dim TipoFornecimento" src="https://github.com/user-attachments/assets/b628cf4f-4ed0-4512-90c0-e2b318b4a6d9" />

 The fifth, merge to Dim Tipo Entidade:

 <img width="845" height="771" alt="Execute Mege tipo entidade" src="https://github.com/user-attachments/assets/3da02f77-fb66-446d-ba33-3d79d525ff49" />

 The sixth, merge to Tipo Pagamento:

 <img width="843" height="769" alt="Execute Merge to Dim Pagamentos" src="https://github.com/user-attachments/assets/09c00592-9173-4fed-9862-0202ad11b292" />

 And finally, the merge to the fact Pagamentos table. 

<img width="839" height="768" alt="Execute Merge to Fact Pagamentos" src="https://github.com/user-attachments/assets/0d3ab6ce-c510-4867-81c9-8ed0ddde0778" />



### Refresh Views Sequence Container

This sequence container, the last one before the tasks focused on sending the email, is, as the previous one, focused on executing sql statements. In this case, simple select statements in the views.  

It contains the following Execute SQL Tasks:

<img width="809" height="365" alt="Update Views Sequence Container 1" src="https://github.com/user-attachments/assets/29e39285-1c03-4149-b489-22848d01e585" />
<img width="812" height="350" alt="Update Views Sequence Container 2" src="https://github.com/user-attachments/assets/484498ad-4a12-4bbb-b5d9-38924aca34b3" />

As the Execute SQL Tasks are simply executing the select statement related to that view and connects to the db, one image will sufice. The SELECT SQL Scripts are in the Select Views file.

<img width="842" height="771" alt="Update View Pagamento em Atraso Agregado" src="https://github.com/user-attachments/assets/7394e032-36ad-4ad1-aeec-61e9122aaa19" />


### Execute SP Generate HTML Report Task

This task executes the store procedure GenerateFinancialHtmlReport. It initalizes the variable, stores the variable with the body resulting from the store procedure in the @Html variable and selects it to store the end result in the ssis variable HtmlReportBody to be passed to the next task. 

The sql script used in this task is:

```sql
SET NOCOUNT ON;
   DECLARE @Html NVARCHAR(MAX);
   EXEC dbo.GenerateFinancialHtmlReport @OutputHTML = @Html OUTPUT;
   SELECT @Html AS HtmlReport;
```
The task was defined as follows:

<img width="840" height="769" alt="Execute SP for HTML_Improved" src="https://github.com/user-attachments/assets/40f4e91e-9dac-427c-b18d-0f83fbc7a8cf" />

The result set was defined as:

<img width="841" height="771" alt="Execute SP for HTML Email result set" src="https://github.com/user-attachments/assets/7fd40309-6715-40e2-a0e4-7db6ea6e11dd" />


### Send Email Script Task

The send email tasks receives the variable with the output generated in the previous task, stores it as string and sends it. This is done through a C# Script.

The variables used are: 

User::HtmlReportBody,$Package::appPass,$Package::ReceivingEmail,$Package::SendingEmail

<img width="839" height="770" alt="Send HTML Email" src="https://github.com/user-attachments/assets/24c296a6-c7b4-4152-a627-fabcce3e1d66" />

The C# Scrip begins by capturing the @HtmlOutput variable result to the C# variable htmlBody:

```C#
using System;
using Microsoft.SqlServer.Dts.Runtime;

namespace ScriptTaskProject
{
    [Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        public void Main()
        {
            try
            {
                // 1. Recupera o HTML que o Parameter Mapping injetou na vari�vel como String pura
                string htmlBody = Dts.Variables["User::HtmlReportBody"].Value.ToString();
```

The next code block assesses the length of the variable htmlBody, reports a message and interrupts the execution, if the variable is empty:

```C#
bool fireAgain = true;
                Dts.Events.FireInformation(0, "Email", "HtmlReportBody length: " + htmlBody.Length, "", 0, ref fireAgain);

                if (string.IsNullOrWhiteSpace(htmlBody))
                {
                    Dts.Events.FireError(0, "Email", "HtmlReportBody is empty when the script runs.", "", 0);
                    Dts.TaskResult = (int)ScriptResults.Failure;
                    return;
                }
```

The next block calls the variables that contain the sending and receiving email addresses and the pass to connect to the the email service provider:

```C#
string fromEmail = Dts.Variables["$Package::SendingEmail"].Value.ToString(); string toEmails = Dts.Variables["$Package::ReceivingEmail"].Value.ToString(); string appPassword = Dts.Variables["$Package::appPass"].GetSensitiveValue()?.ToString();
```

After that, the message is constructed. Firstly, the  sending and receiving email variables are assigned to the .NET class MailMessage to create the mail object and creates the send from and to attributes. 
The second block below builds the body of the email and the necessary attributes to send the email correctly:

```C#
System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage();
                mail.From = new System.Net.Mail.MailAddress(fromEmail);
                mail.To.Add(toEmails);

                mail.Subject = "Relatório de Controlo de Pagamentos";
                mail.SubjectEncoding = System.Text.Encoding.UTF8;
                mail.Body = htmlBody;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                mail.IsBodyHtml = true;
```

After, the SMTP client is configured:

```C#
System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient("smtp.gmail.com", 587);
                smtp.Credentials = new System.Net.NetworkCredential(fromEmail, appPassword);
                smtp.EnableSsl = true;
```

And finally, the email is sent and error handling code is added in case the sending fails:

```C#
smtp.Send(mail);

                Dts.TaskResult = (int)ScriptResults.Success;
            }
            catch (Exception ex)
            {
                Dts.Events.FireError(0, "Erro no Script de Envio de E-mail", ex.Message, string.Empty, 0);
                Dts.TaskResult = (int)ScriptResults.Failure;
            }
        }

        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        };
    }
}
```

### Delete MapaPagamentos Script Task

Finally, the last task, uses a C# script to localize and delete the copied excel sheet.

The ReadOnlyVariables are User::Destination_Excel_MapaPagamentos,User::LoadDate, which contain the filepath and the load date, that is just used as a timestamp.

<img width="842" height="771" alt="Delete MapaPagametos Task" src="https://github.com/user-attachments/assets/ab72d790-bbee-435d-ab64-1916d91e44ae" />

The C# Code, as an overvue, loads the variables to local string variables, has a conditional if statement that deletes the file if exists in the location and with the name provided and shows a message that it was deleted. 
After, it has a simple catch error statement.

```C#
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;



#region Namespaces
using System;
using System.Data;
using Microsoft.SqlServer.Dts.Runtime;
using System.Windows.Forms;
using System.IO;
#endregion

namespace ST_bbe63ef538a74d0e989c48d5cea6903e
{
    /// <summary>
    /// ScriptMain is the entry point class of the script.  Do not change the name, attributes,
    /// or parent of this class.
    /// </summary>
	[Microsoft.SqlServer.Dts.Tasks.ScriptTask.SSISScriptTaskEntryPointAttribute]
    public partial class ScriptMain : Microsoft.SqlServer.Dts.Tasks.ScriptTask.VSTARTScriptObjectModelBase
    {
        private bool _fireAgain = true;

        enum ScriptResults
        {
            Success = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Success,
            Failure = Microsoft.SqlServer.Dts.Runtime.DTSExecResult.Failure
        }




        /// <summary>
        /// This method is called when this script task executes in the control flow.
        /// Before returning from this method, set the value of Dts.TaskResult to indicate success or failure.
        /// To open Help, press F1.
        /// </summary>
        public void Main()
        {
            {
                try
                {
                    string destinationPath = (string)Dts.Variables["User::Destination_Excel_MapaPagamentos"].Value;
                    DateTime loadDate = (DateTime)Dts.Variables["User::LoadDate"].Value;

                    string destFileName = string.Format(
                        "MapaPagamentos2026.xlsx", loadDate
                    );
                    string destPath = Path.Combine(destinationPath, destFileName);

                    if (File.Exists(destPath))
                    {
                        File.Delete(destPath);
                        Dts.Events.FireInformation(
                            0,
                            "Delete MapaPagamentos2026",
                            $"Deleted '{destPath}'.",
                            string.Empty,
                            0,
                            ref _fireAgain);
                    }

                    Dts.TaskResult = (int)ScriptResults.Success;
                }
                catch (Exception ex)
                {
                    Dts.Events.FireError(
                        0,
                        "Delete Excel File",
                        ex.Message,
                        string.Empty,
                        0);

                    Dts.TaskResult = (int)ScriptResults.Failure;
                }
            }
        }
    }
}
```

This concludes the control flow of the Payments Control SSIS Package. 

Just for curiosity, an example of an email sent can be found below:

[Mail - Angel Gutiérrez - Outlook.pdf](https://github.com/user-attachments/files/34500466/Mail.-.Angel.Gutierrez.-.Outlook.pdf)




















