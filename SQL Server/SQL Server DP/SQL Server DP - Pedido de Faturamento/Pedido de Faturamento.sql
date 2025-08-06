USE dbName
GO


SELECT Pedido.NumPedido
     , Cliente.cdcliente
     , Empresa.razao
     , CONVERT(VARCHAR(10), Pedido.DtPedido, 103)                                                      AS 'DtPedido'
     , CONVERT(VARCHAR(10), Pedido.DtFaturamento, 103)                                                 AS 'DtFaturamento'
     , CASE Pedido.StPedido
           WHEN 'A' THEN 'Aberto'
           WHEN 'F' THEN 'Fechado'
           WHEN 'E' THEN 'Encerrado'
           ELSE Pedido.StPedido
    END                                                                                                AS 'Status'
     , Pedido.dsoperacao
     , LTRIM(RTRIM(Vendedor.Nome))                                                                     AS 'Vendedor'
     , CondicaoPgto.descricao
	 , LTRIM(RTRIM(Produto.CdProduto))																   AS 'CdProduto'
     , Produto.DsVenda
     , CAST(PedidoItem.qtdproduto AS INT)                                                              AS 'Qtd'
     , FORMAT(PedidoItem.vlvenda, 'C', 'PT-BR')                                                        AS 'VlVenda'
     , StKit
     , PedidoItem.cdnatoperacao
--     , (select FORMAT(SUM(vlvenda), 'C', 'PT-BR') from PedidoItem where PedidoItem.numpedido IN ('16352', '16353')) AS 'TotalPedido'
FROM Pedido (NOLOCK)
         INNER JOIN Cliente ON Pedido.CdCliente = Cliente.cdcliente
         INNER JOIN Empresa ON Cliente.empresaid = Empresa.empresaid
         INNER JOIN Vendedor ON Pedido.CdVendedor = Vendedor.CdVendedor
         INNER JOIN CondicaoPgto ON Pedido.CdCondPgto = CondicaoPgto.cdcondpgto
         INNER JOIN PedidoItem ON Pedido.NumPedido = PedidoItem.numpedido
         INNER JOIN Produto ON PedidoItem.cdproduto = Produto.CdProduto
WHERE Pedido.NumPedido IN ('167212')