
SELECT cupomitem.cdproduto, cupomitem.alqicms, cupomitem.datacupom, 
ExcecaoICMS.AlqICMS, ExcecaoICMS.AlqReducaoICMS, ROUND(ExcecaoICMS.AlqICMS * (1 - ExcecaoICMS.AlqReducaoICMS / 100), 2)
FROM cupomitem
INNER JOIN pedido ON pedido.numpedido = cupomitem.numpedido
INNER JOIN cliente ON cliente.cdcliente = pedido.CdCliente
INNER JOIN empresa ON empresa.empresaid = cliente.empresaid
INNER JOIN ExcecaoICMS ON ExcecaoICMS.CdProduto = cupomitem.cdproduto AND ExcecaoICMS.uf = empresa.uf

WHERE cupomitem.cdproduto 
IN(
SELECT DISTINCT cdproduto FROM ExcecaoICMS where AlqReducaoICMS > 0)
AND cupomitem.alqicms <> ROUND(ExcecaoICMS.AlqICMS * (1 - ExcecaoICMS.AlqReducaoICMS / 100), 2)
AND YEAR(cupomitem.datacupom) = 2017
ORDER BY cupomitem.datacupom