
--Produto com mais de uma categoria
SELECT CdProduto FROM produto 
WHERE (
SELECT COUNT(produtocategoria.cdproduto)
FROM produtocategoria 
INNER JOIN PlanoCategProduto ON PlanoCategProduto.cdcategoria = produtocategoria.CdCategoria AND PlanoCategProduto.Classe = 'A'
WHERE ProdutoCategoria.CdProduto = produto.cdproduto AND ProdutoCategoria.TpRegistro = produto.TpProduto
) > 1

--Cliente com mais de uma categoria
SELECT cdcliente FROM Cliente
WHERE (
SELECT COUNT(ClienteCategoria.CdCliente)
FROM ClienteCategoria 
INNER JOIN PlanoCategCliente ON PlanoCategCliente.cdcategoria = ClienteCategoria.CdCategoria AND PlanoCategCliente.Classe = 'A'
WHERE ClienteCategoria.cdcliente = Cliente.cdcliente AND ClienteCategoria.EmpresaID = Cliente.empresaid) > 1

