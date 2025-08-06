EXEC msdb.dbo.sp_delete_database_backuphistory @database_name = N'nome do banco de dados'
GO

use [nome do banco de dados];
GO

USE [master]
GO

ALTER DATABASE [nome do banco de dados] SET  SINGLE_USER WITH ROLLBACK IMMEDIATE
GO

USE [master]
GO

DROP DATABASE [nome do banco de dados]
GO
