SELECT PIS_Base_Calculo = ROUND(CASE
                                    WHEN naturezaoperacao.stcalcpis = 'S' THEN
                                                        pendenciavendaitem.vlcalculado *
                                                        pendenciavendaitem.qtdsolicitada
                                                    + CASE
                                                          WHEN naturezaoperacao.StFreteEmBsPIS = 'S' THEN
                                                              pendenciavendaitem.vlfrete * pendenciavendaitem.qtdsolicitada
                                                          ELSE
                                                              0
                                                            END
                                                    + CASE
                                                          WHEN NaturezaOperacao.StDespAcessoriaEmBsPIS = 'S' THEN
                                                              pendenciavendaitem.vldespesa * pendenciavendaitem.qtdsolicitada
                                                          ELSE
                                                              0
                                                            END
                                                - CASE
                                                      WHEN NaturezaOperacao.StDeduzDescontoEmBsPIS = 'S' THEN
                                                          pendenciavendaitem.vldescontoemnf * pendenciavendaitem.qtdsolicitada
                                                      ELSE
                                                          0
                                                            END
                                            + CASE
                                                  WHEN NaturezaOperacao.StImpostoRetidoSTEmBsPIS = 'S' THEN
                                                      pendenciavendaitem.vlimpostoretido
                                                  ELSE
                                                      0
                                                            END
                                            + CASE
                                                  WHEN NaturezaOperacao.StIPIEmBsPIS = 'S' THEN
                                                      pendenciavendaitem.vlipi
                                                  ELSE
                                                      0
                                                            END
                                    ELSE 0 END, 2),

       PIS_Alq_Calculo = CASE
                             WHEN ISNULL(naturezaoperacao.stcalcpis, 'S') = 'S' THEN
                                 CASE
                                     WHEN produto.ncm IS NULL THEN
                                         ddempresa.PIS
                                     ELSE
                                         ISNULL(cencm.alqpis, 0)
                                     END
                             ELSE
                                 0
           END,

       PIS_Valor_Calculo = ROUND(CASE
                                     WHEN naturezaoperacao.stcalcpis = 'S' THEN
                                                         pendenciavendaitem.vlcalculado *
                                                         pendenciavendaitem.qtdsolicitada
                                                     + CASE
                                                           WHEN naturezaoperacao.StFreteEmBsPIS = 'S' THEN
                                                               pendenciavendaitem.vlfrete * pendenciavendaitem.qtdsolicitada
                                                           ELSE
                                                               0
                                                             END
                                                     + CASE
                                                           WHEN NaturezaOperacao.StDespAcessoriaEmBsPIS = 'S' THEN
                                                               pendenciavendaitem.vldespesa * pendenciavendaitem.qtdsolicitada
                                                           ELSE
                                                               0
                                                             END
                                                 - CASE
                                                       WHEN NaturezaOperacao.StDeduzDescontoEmBsPIS = 'S' THEN
                                                           pendenciavendaitem.vldescontoemnf * pendenciavendaitem.qtdsolicitada
                                                       ELSE
                                                           0
                                                             END
                                             + CASE
                                                   WHEN NaturezaOperacao.StImpostoRetidoSTEmBsPIS = 'S' THEN
                                                       pendenciavendaitem.vlimpostoretido
                                                   ELSE
                                                       0
                                                             END
                                             + CASE
                                                   WHEN NaturezaOperacao.StIPIEmBsPIS = 'S' THEN
                                                       pendenciavendaitem.vlipi
                                                   ELSE
                                                       0
                                                             END
                                     ELSE 0 END *
                                 CASE
                                     WHEN ISNULL(naturezaoperacao.stcalcpis, 'S') = 'S' THEN
                                         CASE
                                             WHEN produto.ncm IS NULL THEN
                                                 ddempresa.PIS
                                             ELSE
                                                 ISNULL(cencm.alqpis, 0)
                                             END
                                     ELSE
                                         0
                                     END
                                     / 100, 2),

       COFINS_Base_Calculo = ROUND(CASE
                                       WHEN naturezaoperacao.StCalcCOFINS = 'S' THEN
                                                           pendenciavendaitem.vlcalculado *
                                                           pendenciavendaitem.qtdsolicitada
                                                       + CASE
                                                             WHEN naturezaoperacao.StFreteEmBsCOFINS = 'S' THEN
                                                                 pendenciavendaitem.vlfrete * pendenciavendaitem.qtdsolicitada
                                                             ELSE
                                                                 0
                                                               END
                                                       + CASE
                                                             WHEN NaturezaOperacao.StDespAcessoriaEmBsCOFINS = 'S' THEN
                                                                 pendenciavendaitem.vldespesa * pendenciavendaitem.qtdsolicitada
                                                             ELSE
                                                                 0
                                                               END
                                                   - CASE
                                                         WHEN NaturezaOperacao.StDeduzDescontoEmBsCOFINS = 'S' THEN
                                                             pendenciavendaitem.vldescontoemnf * pendenciavendaitem.qtdsolicitada
                                                         ELSE
                                                             0
                                                               END
                                               + CASE
                                                     WHEN NaturezaOperacao.StImpostoRetidoSTEmBsCOFINS = 'S' THEN
                                                         pendenciavendaitem.vlimpostoretido
                                                     ELSE
                                                         0
                                                               END
                                               + CASE
                                                     WHEN NaturezaOperacao.StIPIEmBsCOFINS = 'S' THEN
                                                         pendenciavendaitem.vlipi
                                                     ELSE
                                                         0
                                                               END
                                       ELSE 0 END, 2),

       COFINS_Alq_Calculo = CASE
                                WHEN ISNULL(naturezaoperacao.StCalcCOFINS, 'S') = 'S' THEN
                                    CASE
                                        WHEN produto.ncm IS NULL THEN
                                            ddempresa.COFINS
                                        ELSE
                                            ISNULL(cencm.alqcofins, 0)
                                        END
                                ELSE
                                    0
           END,

       COFINS_Valor_Calculo = ROUND(CASE
                                        WHEN naturezaoperacao.StCalcCOFINS = 'S' THEN
                                                            pendenciavendaitem.vlcalculado *
                                                            pendenciavendaitem.qtdsolicitada
                                                        + CASE
                                                              WHEN naturezaoperacao.StFreteEmBsCOFINS = 'S' THEN
                                                                  pendenciavendaitem.vlfrete * pendenciavendaitem.qtdsolicitada
                                                              ELSE
                                                                  0
                                                                END
                                                        + CASE
                                                              WHEN NaturezaOperacao.StDespAcessoriaEmBsCOFINS = 'S' THEN
                                                                  pendenciavendaitem.vldespesa * pendenciavendaitem.qtdsolicitada
                                                              ELSE
                                                                  0
                                                                END
                                                    - CASE
                                                          WHEN NaturezaOperacao.StDeduzDescontoEmBsCOFINS = 'S' THEN
                                                              pendenciavendaitem.vldescontoemnf * pendenciavendaitem.qtdsolicitada
                                                          ELSE
                                                              0
                                                                END
                                                + CASE
                                                      WHEN NaturezaOperacao.StImpostoRetidoSTEmBsCOFINS = 'S' THEN
                                                          pendenciavendaitem.vlimpostoretido
                                                      ELSE
                                                          0
                                                                END
                                                + CASE
                                                      WHEN NaturezaOperacao.StIPIEmBsCOFINS = 'S' THEN
                                                          pendenciavendaitem.vlipi
                                                      ELSE
                                                          0
                                                                END
                                        ELSE 0 END *
                                    CASE
                                        WHEN ISNULL(naturezaoperacao.StCalcCOFINS, 'S') = 'S' THEN
                                            CASE
                                                WHEN produto.ncm IS NULL THEN
                                                    ddempresa.COFINS
                                                ELSE
                                                    ISNULL(cencm.alqcofins, 0)
                                                END
                                        ELSE
                                            0
                                        END / 100, 2)

FROM pendenciavendaitem
         INNER JOIN pendenciavenda1
                    ON pendenciavenda1.nummovimento = pendenciavendaitem.nummovimento
         INNER JOIN NaturezaOperacao
                    ON NaturezaOperacao.cdnatoperacao = pendenciavendaitem.cdnatoperacao
         INNER JOIN ddempresa
                    ON 1 = 1
         LEFT JOIN produto
                   ON pendenciavendaitem.cdproduto = produto.cdproduto
                       AND
                      pendenciavendaitem.tpregistro = produto.tpproduto
         LEFT JOIN cencm
                   ON cencm.ncm = produto.ncm
WHERE dtinicio >= '2016-11-01'
  AND (pendenciavendaitem.PIS_Valor_Calculo = 0 OR pendenciavendaitem.COFINS_Valor_Calculo = 0)
  AND (NaturezaOperacao.stcalcpis = 'S' OR NaturezaOperacao.stcalccofins = 'S')

