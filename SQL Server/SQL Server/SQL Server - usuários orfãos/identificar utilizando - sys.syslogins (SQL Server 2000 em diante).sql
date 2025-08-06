SELECT
    A.name AS UserName,
    A.[sid] AS UserSID
FROM
    sys.sysusers A WITH(NOLOCK)
    LEFT JOIN sys.syslogins B WITH(NOLOCK) ON A.[sid] = B.[sid]
WHERE
    A.issqluser = 1 
    AND SUSER_NAME(A.[sid]) IS NULL 
    AND IS_MEMBER('db_owner') = 1 
    AND A.[sid] != 0x00
    AND A.[sid] IS NOT NULL 
    AND ( LEN(A.[sid]) <= 16 ) 
    AND B.[sid] IS NULL
ORDER BY
    A.name