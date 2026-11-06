
## Event Handler

The error handling in this package is merely informative, nut leverages built in ssis variables to store error messages and events in a specific sql server table, for query.

The table is created with the following script: 

```sql
USE [Gestao Financeira]
GO

/****** Object:  Table [dbo].[ETL_ErrorLog]    Script Date: 06/11/2026 21:26:07 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[ETL_ErrorLog](
	[ErrorLogID] [int] IDENTITY(1,1) NOT NULL,
	[RunID] [uniqueidentifier] NULL,
	[Name] [nvarchar](200) NOT NULL,
	[ErrorCode] [int] NULL,
	[ErrorMessage] [nvarchar](max) NOT NULL,
	[ErrorTime] [datetime] NOT NULL,
	[MachineName] [nvarchar](200) NULL,
	[UserName] [nvarchar](200) NULL,
PRIMARY KEY CLUSTERED 
(
	[ErrorLogID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[ETL_ErrorLog] ADD  DEFAULT (getdate()) FOR [ErrorTime]
GO


```

After that, a store procedure is created to capture the system variables from ssis, that are mapped in the execute sql task. You can see the variables below:

<img width="1290" height="853" alt="Event Handler Parameter Mapping" src="https://github.com/user-attachments/assets/c00d667e-c9ce-41c6-8beb-f27755ea1b08" />

The store procedure is defined as:

```sql
CREATE OR ALTER   procedure [dbo].[ErrorCapture] 

@RunID uniqueidentifier,
@Name nvarchar(500),
@ErrorCode nvarchar(100),
@ErrorMessage nvarchar(max),
@MachineName nvarchar(256)

as
begin
set nocount on; 
	insert into [Gestao Financeira].[dbo].[ETL_ErrorLog] 
	( 
	[RunID],[Name],[ErrorCode],[ErrorMessage],[ErrorTime],[MachineName],[UserName]
	)
	values 
	(
	@RunID,@Name,@ErrorCode,@ErrorMessage, getdate() ,@MachineName, HOST_NAME()
	);
end;

GO
```
It performs a simple insert of the captured variable values into in the ETL_ErrorLog table. 

And it is executed in the Execute SQL Task in Event Handlers. 

<img width="840" height="771" alt="Event Handler " src="https://github.com/user-attachments/assets/dc21489b-0851-4dd3-8d58-f9f472460faf" />

