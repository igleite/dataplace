IF (OBJECT_ID('tempdb..#Databases') IS NOT NULL) DROP TABLE #Databases
SELECT IDENTITY(INT, 1,1) AS Id, name
INTO #Databases
FROM sys.sysdatabases
WHERE dbid > 4 -- Ignorar databases de sistema
DECLARE 
@Contador INT = 1, 
@Total_Databases INT = (SELECT COUNT(*) FROM #Databases),
@Query VARCHAR(MAX)
WHILE(@Contador <= @Total_Databases)
BEGIN
SELECT @Query = 'ALTER DATABASE [' + name + '] SET MULTI_USER;'
FROM #Databases
WHERE Id = @Contador
EXEC(@Query)
SET @Contador = @Contador + 1
END