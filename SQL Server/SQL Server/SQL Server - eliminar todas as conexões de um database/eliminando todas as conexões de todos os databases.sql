/*

Caso seja realmente necessário eliminar manualmente as conexões que estão utilizando um database, você pode utilizar esse pequeno trecho de código T-SQL para realizar essa tarefa:

*/

DECLARE @query VARCHAR(MAX) = ''
SELECT @query = COALESCE(@query, ',') + 'KILL ' + CONVERT(VARCHAR, spid) + '; '
FROM master..sysprocesses
WHERE dbid > 4 -- Não eliminar sessões em databases de sistema
  AND spid <> @@SPID -- Não eliminar a sua própria sessão
IF (LEN(@query) > 0)
    PRINT (@query)