DECLARE @Characters Table ([Character] int, [ChangeTo] nchar(1))

INSERT INTO @Characters VALUES (9,   ' ') --Tab, trocar por ' '
INSERT INTO @Characters VALUES (10,  ' ') --Line Feed, trocar por ' '
INSERT INTO @Characters VALUES (13,  ' ') --Carriage Return (Enter), trocar por ' '
INSERT INTO @Characters VALUES (8211,'-' ) --Dash (Unicode Decimal Code &#8211)
INSERT INTO @Characters VALUES (8203,'' ) --Zero Width Space (Unicode Decimal Code &#8203)
INSERT INTO @Characters VALUES (54928,'?' ) --횐
INSERT INTO @Characters VALUES (65411,'?' ) --ﾃ
INSERT INTO @Characters VALUES (65415,'?' ) --ﾇ
INSERT INTO @Characters VALUES (65434,'?' ) --ﾚ

DECLARE @code nvarchar(MAX)
DECLARE @description nvarchar(MAX)

DECLARE @sqlStatement nvarchar(MAX)

DECLARE @table nvarchar(MAX)
DECLARE @fieldTheat nvarchar(MAX)
DECLARE @fieldCondition nvarchar(MAX)
DECLARE @character int
DECLARE @changeTo nchar(1)

--Trocar informações aqui para a execução
--SET @table = 'efdpc_0200'
--SET @fieldTheat = 'descr_item'
--SET @fieldCondition = 'r0200'

SET @table = 'produto'
SET @fieldTheat = 'dscompra'
SET @fieldCondition = 'cdproduto'

SET @sqlStatement = 'DECLARE DeclarationCursor CURSOR FOR 
	SELECT ' + @fieldCondition + ', ' + @fieldTheat + ' 
	FROM ' + @table + ' 
	WHERE 1 = 2'

BEGIN		
		DECLARE Tabela CURSOR FOR SELECT [Character] FROM @Characters		
		OPEN Tabela

		FETCH NEXT FROM Tabela 
		INTO @character
				
		WHILE (@@FETCH_STATUS = 0)
			BEGIN				
				SET @sqlStatement = @sqlStatement + ' OR (' + @fieldTheat + ' LIKE ''%'' + NCHAR(' + CAST(@character AS nvarchar) + ') + ''%'')'
				FETCH NEXT FROM Tabela
				INTO @character			
			END
		CLOSE Tabela
		DEALLOCATE Tabela
	END
		
EXEC sp_executesql @sqlStatement

OPEN DeclarationCursor
 
FETCH NEXT FROM DeclarationCursor
INTO @code, @description
 
WHILE @@FETCH_STATUS = 0
BEGIN  	
	BEGIN		
		DECLARE Tabela CURSOR FOR SELECT [Character], [ChangeTo] FROM @Characters		
		OPEN Tabela

		FETCH NEXT FROM Tabela 
		INTO @character, @changeTo
				
		WHILE (@@FETCH_STATUS = 0)
			BEGIN			
				SET @description = replace(@description, NCHAR(@character), @changeTo)
				
				FETCH NEXT FROM Tabela
				INTO @character, @changeTo			
			END
		CLOSE Tabela
		DEALLOCATE Tabela
	END

	PRINT 'UPDATE ' + @table + ' SET ' + @fieldTheat + ' = ''' + rtrim(@description) + ''' WHERE ' + @fieldCondition + ' = ''' + rtrim(@code) +''''    

	FETCH NEXT FROM DeclarationCursor
    INTO @code, @description
END
CLOSE DeclarationCursor
DEALLOCATE DeclarationCursor;

SET TEXTSIZE 0
SET NOCOUNT ON


