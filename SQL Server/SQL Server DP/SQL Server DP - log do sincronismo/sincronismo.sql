SELECT *
FROM LogEvento (NOLOCK)
WHERE Data BETWEEN '2021-08-03 00:00:00' AND '2021-08-03 23:59:59'
  AND Tipo IN (
               'SyncUpdate', 'SyncInsert', 'SyncDelete', 'SyncStart', 'SyncSuccess', 'SyncConnection', 'SyncError',
               'SyncStDel', 'SyncRollback'
    )
  AND Detalhe IS NOT NULL