DECLARE @cdProduto NVARCHAR(20)

SET @cdProduto = '711572'

SELECT EventoID
     , Data
     , Origem
     , Usuario
     , Computador
     , Tipo
     , Descricao
     , RIGHT(LEFT(Detalhe, 15), 7) AS 'Produto'
     , Detalhe
     , UpdRegistro
FROM LogEvento (NOLOCK)
WHERE Data > '2020-04-08'
  AND Origem LIKE '%fStaPrd01_MNT_ExcessaoICMS%'
--   AND Detalhe LIKE '%' + @cdProduto + '%'