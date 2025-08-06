DECLARE @dbname VARCHAR(MAX)

SET @dbname = 'SPT_IGO_FIXA'

SELECT *
FROM sys.databases
WHERE name LIKE '%' + @dbname + '%'