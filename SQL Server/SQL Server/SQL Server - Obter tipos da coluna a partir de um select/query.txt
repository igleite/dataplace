-- https://stackoverflow.com/questions/1601727/how-do-i-return-the-sql-data-types-from-my-query
-- verifica o datatype que será retornado em cada coluna da query

DECLARE @query nvarchar(max) = N'
SELECT CAST(999.00 AS INT) AS [INT], CAST(999 AS DECIMAL(10,9)) AS [DECIMAL], CAST(999 AS MONEY) AS [MONEY]
'

EXEC sp_describe_first_result_set @query, null, 0;


-- Opção 2

SELECT *
FROM sys.dm_exec_describe_first_result_set
 (
N'
SELECT CAST(999.00 AS INT) AS [INT], CAST(999 AS DECIMAL(10,9)) AS [DECIMAL], CAST(999 AS MONEY) AS [MONEY]
',
NULL,
0
 );


-- com csharp_property mapeando do SQL Server Database Engine type para o .NET Framework type
-- https://learn.microsoft.com/en-us/dotnet/framework/data/adonet/sql-server-data-type-mappings
SELECT
    name,
    is_nullable,
    system_type_id,
    system_type_name,

    CONCAT(
        'public ',
        CASE system_type_id
            -- bigint
            WHEN 127 THEN 'Int64'

            -- binary / varbinary / image / timestamp / rowversion
            WHEN 173 THEN 'Byte[]'   -- binary
            WHEN 165 THEN 'Byte[]'   -- varbinary
            WHEN 34  THEN 'Byte[]'   -- image
            WHEN 189 THEN 'Byte[]'   -- timestamp / rowversion

            -- bit
            WHEN 104 THEN 'Boolean'

            -- char
            WHEN 175 THEN 'String|Char[]'

            -- varchar
            WHEN 167 THEN 'String|Char[]'

            -- nchar
            WHEN 239 THEN 'String|Char[]'

            -- nvarchar
            WHEN 231 THEN 'String|Char[]'

            -- text
            WHEN 35  THEN 'String|Char[]'

            -- ntext
            WHEN 99  THEN 'String|Char[]'

            -- date
            WHEN 40  THEN 'DateTime'

            -- datetime
            WHEN 61  THEN 'DateTime'

            -- datetime2
            WHEN 42  THEN 'DateTime'

            -- smalldatetime
            WHEN 58  THEN 'DateTime'

            -- datetimeoffset
            WHEN 43  THEN 'DateTimeOffset'

            -- time
            WHEN 41  THEN 'TimeSpan'

            -- decimal / numeric
            WHEN 106 THEN 'Decimal'
            WHEN 108 THEN 'Decimal'

            -- money / smallmoney
            WHEN 60  THEN 'Decimal'
            WHEN 122 THEN 'Decimal'

            -- float
            WHEN 62  THEN 'Double'

            -- real
            WHEN 59  THEN 'Single'

            -- int
            WHEN 56  THEN 'Int32'

            -- smallint
            WHEN 52  THEN 'Int16'

            -- tinyint
            WHEN 48  THEN 'Byte'

            -- uniqueidentifier
            WHEN 36  THEN 'Guid'

            -- sql_variant
            WHEN 98  THEN 'Object'

            -- xml
            WHEN 241 THEN 'Xml'

            ELSE 'Object'
        END,
        ' ',
        name,
        ' { get; set; }'
    ) AS csharp_property
FROM sys.dm_exec_describe_first_result_set
(
N'

-- EXEMPLO
SELECT
    -- Inteiros
    CAST(1 AS BIT)                      AS [BIT],
    CAST(1 AS TINYINT)                  AS [TINYINT],
    CAST(1 AS SMALLINT)                 AS [SMALLINT],
    CAST(1 AS INT)                      AS [INT],
    CAST(1 AS BIGINT)                   AS [BIGINT],

    -- Numéricos exatos
    CAST(123.45 AS DECIMAL(10,2))       AS [DECIMAL],
    CAST(123.45 AS NUMERIC(10,2))       AS [NUMERIC],
    CAST(123.45 AS MONEY)               AS [MONEY],
    CAST(123.45 AS SMALLMONEY)          AS [SMALLMONEY],

    -- Numéricos aproximados
    CAST(123.45 AS FLOAT)               AS [FLOAT],
    CAST(123.45 AS REAL)                AS [REAL],

    -- Datas e horas
    CAST(''2026-01-01'' AS DATE)          AS [DATE],
    CAST(''12:34:56'' AS TIME(7))         AS [TIME],
    CAST(''2026-01-01T12:34:56'' AS DATETIME)      AS [DATETIME],
    CAST(''2026-01-01T12:34:56.123'' AS DATETIME2) AS [DATETIME2],
    CAST(''2026-01-01T12:34:56.123-03:00'' AS DATETIMEOFFSET) AS [DATETIMEOFFSET],
    CAST(''2026-01-01'' AS SMALLDATETIME) AS [SMALLDATETIME],

    -- Texto (Unicode e não-Unicode)
    CAST(''ABC'' AS CHAR(10))             AS [CHAR],
    CAST(''ABC'' AS VARCHAR(10))          AS [VARCHAR],
    CAST(''ABC'' AS VARCHAR(MAX))         AS [VARCHAR_MAX],
    CAST(N''ABC'' AS NCHAR(10))            AS [NCHAR],
    CAST(N''ABC'' AS NVARCHAR(10))         AS [NVARCHAR],
    CAST(N''ABC'' AS NVARCHAR(MAX))        AS [NVARCHAR_MAX],
    CAST(''ABC'' AS TEXT)                  AS [TEXT],
    CAST(N''ABC'' AS NTEXT)                AS [NTEXT],

    -- Binários
    CAST(0x010203 AS BINARY(5))          AS [BINARY],
    CAST(0x010203 AS VARBINARY(10))      AS [VARBINARY],
    CAST(0x010203 AS VARBINARY(MAX))     AS [VARBINARY_MAX],
    CAST(''ABC'' AS IMAGE)                 AS [IMAGE],

    -- Outros tipos
    CAST(NEWID() AS UNIQUEIDENTIFIER)   AS [UNIQUEIDENTIFIER],
    CAST(''<x>1</x>'' AS XML)              AS [XML],
    CAST(''test'' AS SYSNAME)              AS [SYSNAME],
    CAST(''123'' AS SQL_VARIANT)           AS [SQL_VARIANT]
',
NULL,
0
);
