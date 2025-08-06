-- Verifica quais usuarios estão usando um banco

USE master
GO
SELECT ec.client_net_address
     , es.[program_name]
     , es.[host_name]
     , es.login_name
     , db.name                        AS 'DatabaseName'
     , ec.connect_time
     , ec.session_id
     , CONCAT('KILL ', ec.session_id) AS 'kill'

FROM sys.dm_exec_sessions AS es
         INNER JOIN sys.dm_exec_connections AS ec ON es.session_id = ec.session_id
         INNER JOIN sys.databases AS db ON es.database_id = db.database_id
WHERE 1 = 1
	AND db.name = '' -- colocar o nome do banco de dados desejado
ORDER BY ec.client_net_address, es.[program_name];
