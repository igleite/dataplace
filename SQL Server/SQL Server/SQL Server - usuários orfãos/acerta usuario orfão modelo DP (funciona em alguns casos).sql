-- Executar a rotina abaixo para cada banco de dados.

DECLARE @usuario AS NVARCHAR(30)
DECLARE @senha AS   NVARCHAR(30)

DECLARE oCursor CURSOR
    FOR
    SELECT username,
           senha
    FROM Usuario

OPEN oCursor

FETCH NEXT FROM oCursor INTO @usuario, @senha

WHILE @@fetch_status = 0
    
    BEGIN
        
        PRINT @usuario + ' - ' + @senha
        EXEC sp_change_users_login 'Auto_Fix', @usuario, null, @senha

        FETCH NEXT FROM oCursor INTO @usuario, @senha
    
    END

CLOSE oCursor
DEALLOCATE oCursor