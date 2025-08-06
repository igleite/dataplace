DECLARE @dbName    SYSNAME
DECLARE @tableName SYSNAME
DECLARE @cmd       VARCHAR(MAX)


DECLARE cursordb CURSOR FOR
    SELECT name AS 'name'
    FROM sys.databases
    WHERE name LIKE 'PNOBRE%'

OPEN cursordb
FETCH NEXT FROM cursordb INTO @dbName

WHILE (@@FETCH_STATUS = 0) BEGIN


    /** cursor 2 */

    DECLARE cursortable CURSOR FOR
        SELECT TABLE_NAME AS 'tbl'
        FROM INFORMATION_SCHEMA.TABLES
        WHERE TABLE_NAME = 'NtFiscal'

    OPEN cursortable
    FETCH NEXT FROM cursortable INTO @tableName

    WHILE (@@FETCH_STATUS = 0) BEGIN


        /** select */

        SET @cmd = 'SELECT ''' + @dbName + '''  as ''Empresa'', * FROM ' + @dbName + '..' + @tableName + ' WHERE NumNota LIKE  ''%28733%'''

        EXEC (@cmd)

        /** end select */


        FETCH NEXT FROM cursortable INTO @tableName

    END

--     PRINT 'FINISHED'
    CLOSE cursortable
    DEALLOCATE cursortable

    /** end cursor 2 */

    FETCH NEXT FROM cursordb INTO @dbName

END

-- PRINT 'FINISHED'
CLOSE cursordb
DEALLOCATE cursordb