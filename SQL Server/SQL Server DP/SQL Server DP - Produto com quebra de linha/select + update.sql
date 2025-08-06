SELECT VAR, CDPRODUTO, DSVENDA
FROM (
         SELECT CASE
                    WHEN var LIKE '%\n%' THEN 'Tem quebra de linha'
             END          AS 'var'
              , CdProduto AS 'CdProduto'
              , DsVenda   AS 'DsVenda'
         FROM (
                  SELECT REPLACE(REPLACE(DsVenda, CHAR(13), ' \n '), CHAR(10),
                                 ' \n ') AS 'var'
                       , CdProduto       AS 'CdProduto'
                       , DsVenda         AS 'DsVenda'
                  FROM Produto
              ) tbl
     ) tbl2
WHERE var IS NOT NULL


UPDATE Produto
SET DsVenda = REPLACE(REPLACE(DsVenda, CHAR(13), ''), CHAR(10), '')