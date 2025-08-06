SELECT cdnatoperacaoOrigem  = Origem.cdnatoperacao
     , cdnatoperacaoDestino = Destino.cdnatoperacao
FROM TURTT_DbTurtt..NaturezaOperacao Origem
         INNER JOIN TURTT_DbCET..NaturezaOperacao Destino
                    ON Origem.cdnatoperacao COLLATE SQL_Latin1_General_CP1_CI_AS = Destino.cdnatoperacao COLLATE SQL_Latin1_General_CP1_CI_AS