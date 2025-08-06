SELECT *
FROM (
         SELECT CdProduto
              , CdCampo
              , Conteudo
         FROM CpoAdicionalProduto
         WHERE CdProduto = '7271'
           AND CdCampo IN ('ALR01', 'ALR03', 'ALR02', 'ALR04', 'ALR05', 'ALR07')
     ) tbaux
         PIVOT (
         MAX(Conteudo)
         FOR CdCampo IN ([ALR01], [ALR03], [ALR02], [ALR04], [ALR05], [ALR07])
         ) AS pvt