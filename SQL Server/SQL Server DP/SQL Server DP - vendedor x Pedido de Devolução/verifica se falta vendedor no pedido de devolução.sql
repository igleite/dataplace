SELECT numpedido
     , cdvendedor
     , NotaOriginalPedido
     , IIF(CountDev <> CountFatOrigem, 'Falta de vendedor no pedido de devolucao', 'OK') AS 'Valida'
     , Dt
FROM (
         SELECT DISTINCT pedidovendedor.numpedido
                       , pedidovendedor.cdvendedor
                       , PedidoItem.NotaOriginalPedido

                       , (SELECT COUNT(DISTINCT cdvendedor)
                          FROM pedidovendedor PP
                          WHERE pedidovendedor.numpedido = PP.numpedido)       AS 'CountDev'

                       , (SELECT COUNT(DISTINCT cdvendedor)
                          FROM pedidovendedor PV
                          WHERE PV.numpedido = PedidoItem.NotaOriginalPedido)  AS 'CountFatOrigem'

                       , CONVERT(VARCHAR(10), pedidovendedor.updregistro, 103) AS 'Dt'
         FROM pedidovendedor
                  INNER JOIN PedidoItem ON pedidovendedor.numpedido = PedidoItem.numpedido
         WHERE pedidovendedor.numpedido IN (
             SELECT Pedido.NumPedido
             FROM Pedido
             WHERE modelopedido = 'DEV'
         )
     ) TbAux
WHERE CountDev <> CountFatOrigem
ORDER BY CountFatOrigem






