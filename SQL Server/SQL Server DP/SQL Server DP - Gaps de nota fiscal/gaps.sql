WITH gaps AS (
    SELECT cur   = NF.nnf
         , nxt   = LEAD(NF.nnf) OVER (ORDER BY NF.nnf)
         , serie = NF.serie
         , mod   = NF.mod
    FROM nfenf NF WITH (NOLOCK)
    WHERE NF.serie = '1'
      AND NF.mod = '55'
      AND NF.nnf >= '3000'
)


SELECT numNota     = NI.numinutilizado
     , dtErrado    = NI.dtcomunicacaosefaz
     , ProxNotaEmi = range.nxt
     , dtCorreto   = dt.demi

     , NI.serie
     , NI.mod

FROM nfenuminutilizado NI WITH (NOLOCK)

         OUTER APPLY (
    SELECT cur   = gaps.cur
         , nxt   = gaps.nxt
         , serie = gaps.serie
         , mod   = gaps.mod
    FROM gaps
) range

         OUTER APPLY (
    SELECT demi = datas.demi
    FROM nfenf datas WITH (NOLOCK)
    WHERE datas.nnf = range.nxt
) dt

WHERE ISNULL(NI.numinutilizado, 0) <> 0
  AND NI.dtcomunicacaosefaz IS NOT NULL
  AND NI.numinutilizado BETWEEN range.cur + 1 AND range.nxt
  AND NI.serie = range.serie
  AND NI.mod = range.mod