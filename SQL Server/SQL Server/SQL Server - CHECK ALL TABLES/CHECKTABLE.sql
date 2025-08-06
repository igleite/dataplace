USE Massas
GO
 
 
DECLARE @tablename sysname
DECLARE @cmd       varchar(MAX)
 
DECLARE crs CURSOR FOR
    SELECT name
    FROM (
             SELECT ROW_NUMBER() OVER (ORDER BY name ASC) AS Row
                  , name                                  AS 'name'
             FROM sys.tables
             WHERE 1 = 1
         ) systables
    WHERE 1 = 1
    ORDER BY name;
 
OPEN crs
 
FETCH NEXT FROM crs INTO @tablename
 
WHILE (@@fetch_status = 0) BEGIN
 
    SET @cmd = 'DBCC CHECKTABLE ( ' + @tablename + ' ) WITH NO_INFOMSGS, ALL_ERRORMSGS, TABLERESULTS;'
    
    PRINT 'INICIANDO DBCC NA TABELA' + @tablename
    EXEC (@cmd)
    PRINT 'FINALIZADO DBCC NA TABELA' + @tablename
 
    FETCH NEXT FROM crs INTO @tablename
 
END
PRINT 'FINISHED'
 
CLOSE crs
 
DEALLOCATE crs