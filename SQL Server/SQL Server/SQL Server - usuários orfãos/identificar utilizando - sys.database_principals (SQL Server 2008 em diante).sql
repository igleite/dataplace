SELECT
    A.name AS UserName,
    A.[sid] AS UserSID
FROM
    sys.database_principals A WITH(NOLOCK)
    LEFT JOIN sys.sql_logins B WITH(NOLOCK) ON A.[sid] = B.[sid]
    JOIN sys.server_principals C WITH(NOLOCK) ON A.[name] COLLATE SQL_Latin1_General_CP1_CI_AI = C.[name] COLLATE SQL_Latin1_General_CP1_CI_AI
WHERE
    A.principal_id > 4
    AND B.[sid] IS NULL
    AND A.is_fixed_role = 0
    AND C.is_fixed_role = 0
    AND A.name NOT LIKE '##MS_%'
    AND A.[type_desc] = 'SQL_USER'
    AND C.[type_desc] = 'SQL_LOGIN'
    AND A.name NOT IN ('sa')
    AND A.authentication_type <> 0 -- NONE
ORDER BY
    A.name