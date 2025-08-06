https://stackoverflow.com/questions/706664/generate-sql-create-scripts-for-existing-tables-with-query

set transaction isolation level read uncommitted;
SET NOCOUNT ON;
SET STATISTICS TIME OFF;
SET STATISTICS IO OFF;

-- http://stackoverflow.com/questions/12639948/sql-nvarchar-and-varchar-limits
--NB: Crazy table can still have truncation at 4000 because of unexpected number of indexes or other very long list of columns/defaults etc
-- triggers are not supported
-- xml indexes are not supported

DECLARE @Tables table(id int identity(1,1), [name] sysname);
insert into @Tables([name])
values
   ('<yourSchema>.<youTableName>')
  ,('<yourSchema2>.<youTableName2>')
;
DECLARE @object_id                          int;
DECLARE @SourceDatabase nvarchar(max) = N'SourceTest'; --this is used only by the insert 
DECLARE @TargetDatabase nvarchar(max) = N'DescTest';   --this is used only by the insert and USE <DBName>

--- options ---
DECLARE @UseTransaction                     bit = 0; 
DECLARE @GenerateUseDatabase                bit = 0;
DECLARE @GenerateFKs                        bit = 1;
DECLARE @GenerateIdentity                   bit = 1;
DECLARE @GenerateCollation                  bit = 0;
DECLARE @GenerateCreateTable                bit = 1;
DECLARE @GenerateIndexes                    bit = 1;
DECLARE @GenerateConstraints                bit = 1;
DECLARE @GenerateKeyConstraints             bit = 1;
DECLARE @GenerateConstraintNameOfDefaults   bit = 1;
DECLARE @GenerateDropIfItExists             bit = 1;
DECLARE @GenerateDropFKIfItExists           bit = 0;
DECLARE @GenerateDelete                     bit = 0;
DECLARE @GenerateInsertInto                 bit = 0;
DECLARE @GenerateIdentityInsert             int = 0; --0 ignore set,but add column; 1 generate; 2 ignore set and column
DECLARE @GenerateSetNoCount                 int = 2; --0 ignore set,1=set on, 2=set off 
DECLARE @GenerateMessages                   bit = 1; --print with no wait
DECLARE @GenerateDataCompressionOptions     bit = 1; --TODO: generates the compression option only of the table, not the indexes
                                                    --NB: the compression options reflects the design value.
                                                    --The actual compression of a the page is saved here
                                                    --SELECT * from sys.dm_db_database_page_allocations(DB_ID(), @object_ID, 0, 1, 'DETAILED')

-----------------------------------------------------------------------------
------------------------------------------------------------------------------
--- Let's play
DECLARE @DataTypeSpacer                     int = 45; --this is just to improve the formatting of the script ...
DECLARE @name                               sysname;
DECLARE @SQL                                NVARCHAR(MAX) = N''

DECLARE db_cursor CURSOR FOR SELECT [name] from @Tables
OPEN db_cursor  
FETCH NEXT FROM db_cursor INTO @name  

WHILE @@FETCH_STATUS = 0  
BEGIN  
    SET @object_id = object_ID(@name)
    goto CreateScript;
backFromCreateScript:
    FETCH NEXT FROM db_cursor INTO @name 
END 
CLOSE db_cursor  
DEALLOCATE db_cursor 
return;

CreateScript:

DECLARE @CR NVARCHAR(max) = NCHAR(13);
DECLARE @TB NVARCHAR(max) = NCHAR(9);
DECLARE @CurrentIndent nvarchar(max) = ''

;WITH index_column AS 
(
    SELECT 
        ic.[object_id]
        , ic.index_id
        , ic.is_descending_key
        , ic.is_included_column
        , c.name
    FROM sys.index_columns ic WITH (NOWAIT)
    JOIN sys.columns c WITH (NOWAIT) ON ic.[object_id] = c.[object_id] AND ic.column_id = c.column_id
    WHERE ic.[object_id] = @object_id
),
fk_columns AS 
(
    SELECT 
        k.constraint_object_id
        , cname = c.name
        , rcname = rc.name
    FROM sys.foreign_key_columns k WITH (NOWAIT)
    JOIN sys.columns rc WITH (NOWAIT) ON rc.[object_id] = k.referenced_object_id AND rc.column_id = k.referenced_column_id 
    JOIN sys.columns c WITH (NOWAIT) ON c.[object_id] = k.parent_object_id AND c.column_id = k.parent_column_id
    WHERE k.parent_object_id = @object_id and @GenerateFKs = 1
)
SELECT @SQL = 

    --------------------  USE DATABASE   --------------------------------------------------------------------------------------------------
        CAST(
            CASE WHEN @GenerateUseDatabase = 1
            THEN N'USE ' + @TargetDatabase + N';' + @CR
            ELSE N'' END 
        as nvarchar(max))
        +
    --------------------  SET NOCOUNT   --------------------------------------------------------------------------------------------------
        CAST(
            CASE @GenerateSetNoCount 
            WHEN 1 THEN N'SET NOCOUNT ON;' + @CR
            WHEN 2 THEN N'SET NOCOUNT OFF;' + @CR
            ELSE N'' END 
        as nvarchar(max))
        +
    --------------------  USE TRANSACTION  --------------------------------------------------------------------------------------------------
        CAST(
            CASE WHEN @UseTransaction = 1
            THEN 
                N'SET XACT_ABORT ON' + @CR
                + N'BEGIN TRY' + @CR
                + N'BEGIN TRAN' + @CR
            ELSE N'' END 
        as nvarchar(max))
        +
    --------------------  DROP SYNONYM   --------------------------------------------------------------------------------------------------
        CASE WHEN @GenerateDropIfItExists = 1
        THEN 
            CAST(
                    N'IF OBJECT_ID(''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''',''SN'') IS NOT NULL DROP SYNONYM ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N';' + @CR 
            as nvarchar(max))
        ELSE 
            CAST(
                    N'' 
            as nvarchar(max))
        END 
        +
    --------------------  DROP IS Exists --------------------------------------------------------------------------------------------------
        CASE WHEN @GenerateDropIfItExists = 1
        THEN 
            --Drop table if exists
            CAST(
                N'IF OBJECT_ID(''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''',''U'') IS NOT NULL DROP TABLE ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N';' + @CR 
            as nvarchar(max))
            + @CR
        ELSE N'' END 
        +
    --------------------  DROP IS Exists --------------------------------------------------------------------------------------------------
        CAST((CASE WHEN @GenerateMessages = 1 and @GenerateDropFKIfItExists = 1 THEN 
            N'RAISERROR(''DROP CONSTRAINTS OF %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR            
        ELSE N'' END) as nvarchar(max)) 
        +
        CASE WHEN @GenerateDropFKIfItExists = 1
        THEN 
            --Drop foreign keys
            ISNULL(((
                SELECT 
                    CAST(
                        N'ALTER TABLE ' + quotename(s.name) + N'.' + quotename(t.name) + N' DROP CONSTRAINT ' + RTRIM(f.name) + N';' + @CR
                    as nvarchar(max))
                FROM sys.tables t
                INNER JOIN sys.foreign_keys f ON f.parent_object_id = t.object_id
                INNER JOIN sys.schemas      s ON s.schema_id = f.schema_id
                WHERE f.referenced_object_id = @object_id
                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'))
            ,N'') + @CR
        ELSE N'' END 
    +
    --------------------- CREATE TABLE -----------------------------------------------------------------------------------------------------------------
    CAST((CASE WHEN @GenerateMessages = 1 THEN 
        N'RAISERROR(''CREATE TABLE %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR           
    ELSE N'' END) as nvarchar(max)) 
    +
    CASE WHEN @GenerateCreateTable = 1 THEN 
        CAST(
            N'CREATE TABLE ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + @CR + N'(' + @CR + STUFF((
            SELECT 
                CAST(
                    @TB + N',' + quotename(c.name) + N' ' + ISNULL(replicate(' ',@DataTypeSpacer - len(quotename(c.name))),'')  --isnull(replicate) then len(quotename(c.name)) > @DataTypeSpacer
                    +  
                    CASE WHEN c.is_computed = 1
                        THEN N' AS ' + cc.[definition] 
                        ELSE UPPER(tp.name) + 
                            CASE WHEN tp.name IN (N'varchar', N'char', N'varbinary', N'binary', N'text')
                                    THEN N'(' + CASE WHEN c.max_length = -1 THEN N'MAX' ELSE CAST(c.max_length AS NVARCHAR(5)) END + N')'
                                    WHEN tp.name IN (N'nvarchar', N'nchar', N'ntext')
                                    THEN N'(' + CASE WHEN c.max_length = -1 THEN N'MAX' ELSE CAST(c.max_length / 2 AS NVARCHAR(5)) END + N')'
                                    WHEN tp.name IN (N'datetime2', N'time2', N'datetimeoffset') 
                                    THEN N'(' + CAST(c.scale AS NVARCHAR(5)) + N')'
                                    WHEN tp.name = N'decimal' 
                                    THEN N'(' + CAST(c.[precision] AS NVARCHAR(5)) + N',' + CAST(c.scale AS NVARCHAR(5)) + N')'
                                ELSE N''
                            END +
                            CASE WHEN c.collation_name IS NOT NULL and @GenerateCollation = 1 THEN N' COLLATE ' + c.collation_name ELSE N'' END +
                            CASE WHEN c.is_nullable = 1 THEN N' NULL' ELSE N' NOT NULL' END +
                            CASE WHEN dc.[definition] IS NOT NULL THEN CASE WHEN @GenerateConstraintNameOfDefaults = 1 THEN N' CONSTRAINT ' + quotename(dc.name) ELSE N'' END + N' DEFAULT' + dc.[definition] ELSE N'' END + 
                            CASE WHEN ic.is_identity = 1 and @GenerateIdentity = 1 THEN N' IDENTITY(' + CAST(ISNULL(ic.seed_value, N'0') AS NCHAR(1)) + N',' + CAST(ISNULL(ic.increment_value, N'1') AS NCHAR(1)) + N')' ELSE N'' END 
                    END + @CR
                AS nvarchar(Max))
            FROM sys.columns c WITH (NOWAIT)
                INNER JOIN sys.types tp WITH (NOWAIT) ON c.user_type_id = tp.user_type_id
                LEFT JOIN sys.computed_columns cc WITH (NOWAIT) ON c.[object_id] = cc.[object_id] AND c.column_id = cc.column_id
                LEFT JOIN sys.default_constraints dc WITH (NOWAIT) ON c.default_object_id != 0 AND c.[object_id] = dc.parent_object_id AND c.column_id = dc.parent_column_id
                LEFT JOIN sys.identity_columns ic WITH (NOWAIT) ON c.is_identity = 1 AND c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
            WHERE c.[object_id] = @object_id
            ORDER BY c.column_id
            FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, @TB + N' ')
    as nvarchar(max))
    ELSE 
        CAST('' as nvarchar(max)) 
    end 
    + 
    ---------------------- Key Constraints ----------------------------------------------------------------
    CAST(
        case when @GenerateKeyConstraints <> 1 THEN N'' ELSE 
            ISNULL((SELECT @TB + N', CONSTRAINT ' + quotename(k.name) + N' PRIMARY KEY ' + ISNULL(kidx.type_desc, N'') + N'(' + 
                        (SELECT STUFF((
                            SELECT N', ' + quotename(c.name) + N' ' + CASE WHEN ic.is_descending_key = 1 THEN N'DESC' ELSE N'ASC' END
                            FROM sys.index_columns ic WITH (NOWAIT)
                            JOIN sys.columns c WITH (NOWAIT) ON c.[object_id] = ic.[object_id] AND c.column_id = ic.column_id
                            WHERE ic.is_included_column = 0
                                AND ic.[object_id] = k.parent_object_id 
                                AND ic.index_id = k.unique_index_id     
                            FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N''))
                + N')' + @CR
                FROM sys.key_constraints k WITH (NOWAIT) LEFT JOIN sys.indexes kidx ON
                    k.parent_object_id = kidx.object_id and k.unique_index_id = kidx.index_id
                WHERE k.parent_object_id = @object_id 
                    AND k.[type] = N'PK'), N'') + N')'  + @CR
        END 
    as nvarchar(max))
    +
    CAST(
    CASE 
        WHEN 
            @GenerateDataCompressionOptions = 1 
            AND 
            (SELECT top 1 data_compression_desc from sys.partitions where object_ID = @object_id and index_id = 1) <> N'NONE'
        THEN 
            N'WITH (DATA_COMPRESSION=' + (SELECT top 1 data_compression_desc from sys.partitions where object_ID = @object_id and index_id = 1) + N')' + @CR
        ELSE
            N'' + @CR
    END as nvarchar(max))
    + 
    --------------------- FOREIGN KEYS -----------------------------------------------------------------------------------------------------------------
    CAST((CASE WHEN @GenerateMessages = 1 and @GenerateDropFKIfItExists = 1 THEN 
        N'RAISERROR(''CREATING FK OF  %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR            
    ELSE N'' END) as nvarchar(max)) 
    +
    CAST(
        ISNULL((SELECT (
            SELECT @CR +
            N'ALTER TABLE ' + + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + + N' WITH' 
            + CASE WHEN fk.is_not_trusted = 1 
                THEN N' NOCHECK' 
                ELSE N' CHECK' 
            END + 
            N' ADD CONSTRAINT ' + quotename(fk.name)  + N' FOREIGN KEY(' 
            + STUFF((
                SELECT N', ' + quotename(k.cname) + N''
                FROM fk_columns k
                WHERE k.constraint_object_id = fk.[object_id]
                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'')
            + N')' +
            N' REFERENCES ' + quotename(SCHEMA_NAME(ro.[schema_id])) + N'.' + quotename(ro.name) + N' ('
            + STUFF((
                SELECT N', ' + quotename(k.rcname) + N''
                FROM fk_columns k
                WHERE k.constraint_object_id = fk.[object_id]
                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'')
            + N')'
            + CASE 
                WHEN fk.delete_referential_action = 1 THEN N' ON DELETE CASCADE' 
                WHEN fk.delete_referential_action = 2 THEN N' ON DELETE SET NULL'
                WHEN fk.delete_referential_action = 3 THEN N' ON DELETE SET DEFAULT' 
                ELSE N'' 
            END
            + CASE 
                WHEN fk.update_referential_action = 1 THEN N' ON UPDATE CASCADE'
                WHEN fk.update_referential_action = 2 THEN N' ON UPDATE SET NULL'
                WHEN fk.update_referential_action = 3 THEN N' ON UPDATE SET DEFAULT'  
                ELSE N'' 
            END 
            + @CR + N'ALTER TABLE ' + + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + + N' CHECK CONSTRAINT ' + quotename(fk.name)  + N'' + @CR
        FROM sys.foreign_keys fk WITH (NOWAIT)
        JOIN sys.objects ro WITH (NOWAIT) ON ro.[object_id] = fk.referenced_object_id
        WHERE fk.parent_object_id = @object_id
        FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)')), N'')
    as nvarchar(max))
    + 
    --------------------- INDEXES ----------------------------------------------------------------------------------------------------------
    CAST((CASE WHEN @GenerateMessages = 1 and @GenerateIndexes = 1 THEN 
        N'RAISERROR(''CREATING INDEXES OF  %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR           
    ELSE N'' END) as nvarchar(max)) 
    +
    case when @GenerateIndexes = 1 THEN 
        CAST(
            ISNULL(((SELECT
                @CR + N'CREATE' + CASE WHEN i.is_unique = 1 THEN N' UNIQUE ' ELSE N' ' END 
                        + i.type_desc + N' INDEX ' + quotename(i.name) + N' ON ' + + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + + N' (' +
                        STUFF((
                        SELECT N', ' + quotename(c.name) + N'' + CASE WHEN c.is_descending_key = 1 THEN N' DESC' ELSE N' ASC' END
                        FROM index_column c
                        WHERE c.is_included_column = 0
                            AND c.index_id = i.index_id
                        FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'') + N')'  
                        + ISNULL(@CR + N'INCLUDE (' + 
                            STUFF((
                            SELECT N', ' + quotename(c.name) + N''
                            FROM index_column c
                            WHERE c.is_included_column = 1
                                AND c.index_id = i.index_id
                            FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)'), 1, 2, N'') + N')', N'')  + @CR
                FROM sys.indexes i WITH (NOWAIT)
                WHERE i.[object_id] = @object_id
                    AND i.is_primary_key = 0
                    AND i.[type] in (1,2)
                    and @GenerateIndexes = 1
                FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)')
            ), N'')
        as nvarchar(max))
    ELSE 
        CAST(N'' as nvarchar(max))
    END 
    +
    ------------------------  @GenerateDelete     ----------------------------------------------------------
    CAST((CASE WHEN @GenerateMessages = 1 and @GenerateDelete = 1 THEN 
        N'RAISERROR(''TRUNCATING  %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR            
    ELSE N'' END) as nvarchar(max)) 
    +
    CASE WHEN @GenerateDelete = 1 THEN
        CAST(
            (CASE WHEN exists (SELECT * FROM sys.foreign_keys WHERE referenced_object_id = @object_id) THEN 
                N'DELETE FROM ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N';' + @CR
            ELSE
                N'TRUNCATE TABLE ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N';' + @CR
            END)
        AS NVARCHAR(max))
    ELSE 
        CAST(N'' as nvarchar(max))
    END 
    +
    ------------------------- @GenerateInsertInto ----------------------------------------------------------
    CAST((CASE WHEN @GenerateMessages = 1 and @GenerateDropFKIfItExists = 1 THEN 
        N'RAISERROR(''INSERTING INTO  %s'',10,1, ''' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N''') WITH NOWAIT;' + @CR            
    ELSE N'' END) as nvarchar(max)) 
    +
    CASE WHEN @GenerateInsertInto = 1
    THEN 
        CAST(
                CASE WHEN EXISTS (SELECT * from sys.columns c where c.[object_id] = @object_id and is_identity = 1) AND @GenerateIdentityInsert = 1 THEN 
                    N'SET IDENTITY_INSERT ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N' ON;' + @CR
                ELSE 
                    CAST('' AS nvarchar(max))
                END 
                +
                N'INSERT INTO ' + QUOTENAME(@TargetDatabase) + N'.' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N'(' 
                + @CR
                +
                (
                    @TB + N' ' + SUBSTRING(
                        (
                        SELECT @TB + ','+ quotename(Name) + @CR 
                        from sys.columns c 
                        where 
                                c.[object_id] = @object_id 
                            AND system_type_ID <> 189 /*timestamp*/ 
                            AND is_computed = 0
                            and (is_identity = 0 or @GenerateIdentityInsert in (0,1))
                        FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)')
                    ,3,99999)

                )
                + N')' + @CR + N'SELECT ' 
                + @CR
                +
                (
                    @TB + N' ' + SUBSTRING(
                        (
                        SELECT @TB + ','+ quotename(Name) + @CR 
                        FROM sys.columns c 
                        WHERE   c.[object_id] = @object_id 
                            and system_type_ID <> 189 /*timestamp*/ 
                            and is_computed = 0                     
                            and (is_identity = 0 or @GenerateIdentityInsert  in (0,1))
                        FOR XML PATH(N''), TYPE).value(N'.', N'NVARCHAR(MAX)')
                    ,3,99999)
                )
                + N'FROM ' + @SourceDatabase +  N'.' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id))            
                + N';' + @CR
                + CASE WHEN EXISTS (SELECT * from sys.columns c where c.[object_id] = @object_id and is_identity = 1)  AND @GenerateIdentityInsert = 1 THEN 
                    N'SET IDENTITY_INSERT ' + quotename(OBJECT_schema_name(@object_id)) + N'.' + quotename(OBJECT_NAME(@object_id)) + N' OFF;'+ @CR
                ELSE 
                    CAST('' AS nvarchar(max))
                END              
        as nvarchar(max))
    ELSE 
        CAST(
                N'' 
        as nvarchar(max))
    END 
    +
    --------------------  USE TRANSACTION  --------------------------------------------------------------------------------------------------
    CAST(
        CASE WHEN @UseTransaction = 1
        THEN 
            @CR + N'COMMIT TRAN; '
            + @CR + N'END TRY'
            + @CR + N'BEGIN CATCH'
            + @CR + N'  IF XACT_STATE() IN (-1,1)'
            + @CR + N'      ROLLBACK TRAN;'
            + @CR + N''
            + @CR + N'  SELECT   ERROR_NUMBER() AS ErrorNumber  '
            + @CR + N'          ,ERROR_SEVERITY() AS ErrorSeverity  '
            + @CR + N'          ,ERROR_STATE() AS ErrorState  '
            + @CR + N'          ,ERROR_PROCEDURE() AS ErrorProcedure  '
            + @CR + N'          ,ERROR_LINE() AS ErrorLine  '
            + @CR + N'          ,ERROR_MESSAGE() AS ErrorMessage; '
            + @CR + N'END CATCH'
        ELSE N'' END 
    as nvarchar(max))

--print is limited to 4000 chars, if necessary, I use multiple print
--to maintain the consistency of the script, I split near the closest CrLF to the max chunk size
DECLARE @i  int = 1;
DECLARE @maxChunk integer = 3990;
DECLARE @len integer = @maxChunk;

WHILE @i < len(@SQL)
BEGIN 
    IF len(@SQL) > (@i + @len)
        set @len = len(substring(@SQL, @i, @maxChunk)) - CHARINDEX(@CR, reverse(substring(@SQL, @i, @len))) + 1
    PRINT substring(@SQL, @i, @len)
    set @i      =  @i + @len
    set @len    =  @maxChunk
END

--SELECT datalength(@SQL), @sql
--EXEC sys.sp_executesql @SQL

goto backFromCreateScript;