SELECT B.session_id                                                                Sessao,
       D.name                                                                      Banco,
       B.login_name                                                                Usuario,
       B.program_name                                                              App,
       B.status                                                                    Status,
       B.host_name                                                                 Maquina,
       CASE WHEN B.session_id = A.blocking_session_id THEN 'Sim' ELSE 'Nao' END AS Responsavel,
       A.blocking_session_id                                                       Causador,
       F.transaction_begin_time                                                    Inicio,
       A.wait_duration_ms                                                          Tempo_ms,
       C.command                                                                   Comando,
       C.cpu_time                                                                  CPU,
       B.memory_usage                                                              MEM
FROM sys.dm_exec_sessions B
         LEFT Join sys.dm_exec_requests C on B.session_id = C.session_id
         LEFT Join sys.databases D on C.database_id = D.database_id
         LEFT Join sys.dm_os_waiting_tasks A on A.session_id = B.session_id
         LEFT JOIN SYS.dm_tran_session_transactions E ON A.session_id = E.session_id
         LEFT JOIN SYS.dm_tran_active_transactions F ON E.TRANSACTION_ID = F.TRANSACTION_ID
Where b.session_id > 50
  AND b.session_id <> @@SPID
GROUP BY B.session_id, D.name, B.login_name, B.program_name, B.status, B.host_name, F.transaction_begin_time,
         CASE WHEN A.session_id = A.blocking_session_id THEN 'Sim' ELSE 'Nao' END, A.blocking_session_id,
         A.wait_duration_ms, C.command, C.cpu_time, B.memory_usage
ORDER BY ISNULL(A.wait_duration_ms, 0) DESC,
         CASE WHEN A.session_id = A.blocking_session_id THEN 'Sim' ELSE 'Nao' END DESC, A.blocking_session_id,
         b.session_id, B.status