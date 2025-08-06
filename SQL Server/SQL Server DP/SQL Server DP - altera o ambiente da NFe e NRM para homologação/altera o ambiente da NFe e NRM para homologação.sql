UPDATE prmnfe
SET wsconsultacadastrocontribuinte = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx',
    wsstatusservico                = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx',
    wsrecepcao                     = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx',
    wsretrecepcao                  = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx',
    wsconsultanfe                  = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
    wsinutilizacaonfe              = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx',
    wsrecepcaoevento               = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx',
    wsconsultanfe_nfce             = N'https://homologacao.nfce.fazenda.sp.gov.br/ws/NFeConsultaProtocolo4.asmx';




UPDATE prmnfe
SET pathschemas                    = N'C:\Symphony\Nfe\Schemas',
    patherro                       = N'C:\Symphony\Nfe\ERRO',
    pathxml                        = N'C:\Symphony\Nfe\XML',
    pathxml_assinado               = N'C:\Symphony\Nfe\XML_Assinado',
    pathxml_transmitido            = N'C:\Symphony\Nfe\XML_Transmitido',
    pathxml_processado             = N'C:\Symphony\Nfe\XML_processado',
    wsconsultacadastrocontribuinte = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/cadconsultacadastro4.asmx',
    wsstatusservico                = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfestatusservico4.asmx',
    wsrecepcao                     = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeautorizacao4.asmx',
    wsretrecepcao                  = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nferetautorizacao4.asmx',
    wsconsultanfe                  = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
    wscancelamentonfe              = N'',
    wsinutilizacaonfe              = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nfeinutilizacao4.asmx',
    wsrecepcaoevento               = N'https://homologacao.nfe.fazenda.sp.gov.br/ws/nferecepcaoevento4.asmx',
    linkconsultanfe                = N'https://nfe.fazenda.sp.gov.br/nfeweb/services/nfeconsulta2.asmx',
    tpambiente                     = N'2',
    tpimpressaodanfe               = N'1';