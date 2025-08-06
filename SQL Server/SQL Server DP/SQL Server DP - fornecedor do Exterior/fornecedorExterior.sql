SELECT LTRIM(RTRIM(F.cdfornecedor)) AS 'cdfornecedor'
     , E.empresaid
     , LTRIM(RTRIM(E.fantasia))     AS 'fantasia'
     , LTRIM(RTRIM(E.razao))        AS 'razao'
     , E.tpcadastro
     , E.endereco
     , E.complemento
     , E.bairro
     , E.cep
     , E.cidade
     , E.uf
     , E.cdpais
     , E.inscriest                  AS 'Estadual'
     , E.inscrifed                  AS 'C.N.P.J.'
     , E.inscrimunic                AS 'Municipal'
     , CASE F.stativocf
           WHEN '1' THEN 'Ativo'
           ELSE 'Inativo'
    END                                'Status'
FROM Empresa E (NOLOCK)
         INNER JOIN Fornecedor F (NOLOCK)
                    ON E.empresaid = F.empresaid
WHERE E.cdpais <> 'BR';




