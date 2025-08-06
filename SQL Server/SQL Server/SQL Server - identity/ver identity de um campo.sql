SELECT sysTables.name                                 AS 'table_name'
     , sysIdentity.name                               AS 'column_name'
     , sysIdentity.is_identity                        AS 'is_identity'
     , sysIdentity.last_value                         AS 'last_value'
     , (CONVERT(NUMERIC, sysIdentity.last_value) + 1) AS 'next_value'
FROM sys.identity_columns sysIdentity
         INNER JOIN sys.tables sysTables ON sysIdentity.object_id = sysTables.object_id
WHERE sysIdentity.name = 'nummovimento'


-- redefinir identity

DBCC CHECKIDENT ( pendenciavenda1, RESEED, 1441)