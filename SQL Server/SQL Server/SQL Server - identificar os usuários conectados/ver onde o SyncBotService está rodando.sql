-- Verifica em qual banco o SyncBotService está executando

USE master
GO
SELECT ec.client_net_address,
       es.[program_name],
       es.[host_name],
       es.login_name,
       db.name AS 'DatabaseName',
       ec.connect_time

FROM sys.dm_exec_sessions AS es
         INNER JOIN sys.dm_exec_connections AS ec ON es.session_id = ec.session_id
         INNER JOIN sys.databases AS db ON es.database_id = db.database_id
WHERE es.[program_name] = 'SyncBotService'
ORDER BY ec.client_net_address, es.[program_name];
