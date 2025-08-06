SELECT produtopreco.cdtabela,
       produtopreco.sqtabela,
       produtopreco.tpitem,
       produtopreco.cdproduto,
       produto.DsVenda,
       produtopreco.vlcusto,
       produtopreco.vlvenda,
       produtopreco.cdindicecusto,
       produtopreco.cdindicevenda,
       produtopreco.obs,
       produto.ValorCustoR
FROM ProdutoPreco
         LEFT JOIN produto ON produtopreco.cdproduto = produto.cdproduto
         LEFT JOIN servico ON produtopreco.cdproduto = servico.cdservico AND produtopreco.tpitem = '5'
WHERE cdtabela = '001'
  AND sqtabela = 1
  AND produtopreco.tpitem = '1'