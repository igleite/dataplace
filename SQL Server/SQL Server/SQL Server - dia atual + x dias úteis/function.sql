-- https://social.technet.microsoft.com/wiki/pt-br/contents/articles/26290.adicionando-dias-uteis-usando-t-sql.aspx

CREATE FUNCTION dbo.ufn_ADD_WORKING_DAYS(
    @DATE DATE,
    @NDAYS INT
) RETURNS DATE
BEGIN

    IF @DATE IS NULL
        BEGIN
            SET @DATE = GETDATE();
        END

    DECLARE @STARTDATE INT = 0
    DECLARE @COUNT INT = 0
    DECLARE @NEWDATE DATE = DATEADD(DAY, 1, @DATE)

    WHILE @COUNT < @NDAYS
        BEGIN
            IF DATEPART(WEEKDAY, @NEWDATE) NOT IN (7, 1) --AND @NEWDATE NOT IN ( SELECT DT_HOLIDAY FROM TB_HOLIDAYS ) 
                SET @COUNT += 1;
            SELECT @NEWDATE = DATEADD(DAY, 1, @NEWDATE), @STARTDATE += 1;
        END

    RETURN DATEADD(DAY, @STARTDATE, @DATE);
END
GO


SELECT dbo.ufn_ADD_WORKING_DAYS(GETDATE(), 5)
GO

