DECLARE @numNf NVARCHAR(50)
DECLARE @ano NVARCHAR(4)
DECLARE @mes NVARCHAR(2)

SET @numNf = '70463'
SET @ano = '2020'
SET @mes = '3'

SELECT LogEvento.Usuario           AS 'Usuario'
     , LogEvento.Data              AS 'Data'
     , LogEvento.Descricao         AS 'Descricao'
     , LogEvento.Detalhe           AS 'Detalhe'
     , LEFT(LogEvento.Detalhe, 63) AS 'ChavedeAcesso'
     , LogEvento.*
FROM LogEvento
WHERE YEAR(LogEvento.Data) >= @ano
  AND MONTH(LogEvento.Data) >= @mes
  AND LogEvento.Origem LIKE '%NFe%'
  AND LogEvento.Descricao LIKE '%Exclusão de NF-e/NFC-e%'
  AND LogEvento.Detalhe LIKE '%' + @numNf + '%'