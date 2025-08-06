DECLARE @chvAcesso VARCHAR(18)
DECLARE @newcod VARCHAR(6)
DECLARE @cep    VARCHAR(8)

SET @chvAcesso = ''
SET @newcod = (SELECT LEFT(NEWID(), 6))
SET @cep = (SELECT COUNT(CEP.cep) FROM CEP WHERE CEP.cep = '17400000')


UPDATE DdEmpresa
SET Razao        = N'DP INFORMÁTICA LTDA - EPP',
    NomeFantasia = N'DATAPLACE',
    DDD          = N'14',
    Fone1        = N'34079000',
    CdPais       = N'BR',
    UF           = N'SP',
    CEP          = N'17400000 ',
    Endereco     = N'R. Cel. Joaquim Piza',
    Numero       = N'358',
    Bairro       = N'Centro',
    Cidade       = N'Garça',
    Chave_Local  = @newcod,
    Site         = N'www.dataplace.com.br',
    TpInscricao  = N'1',
    DtInscricao  = N'2010-07-01 00:00:00.000',
    InscriFed    = N'63917876000117',
    InscriEst    = N'315017592118',
    InscriMunic  = N'1131490',
    TpEmpresa    = N'3',
    CdNaDp       = 1,
    Chave        = @chvAcesso;

INSERT INTO CEP_Localidade
( chave_local
, cdpais
, cdempresa
, cdfilial
, nome_local
, uf
, tipo_local
, cdcidadeibge
, cdCidadeSIAFI
, updregistro
, stdel)
VALUES ( @newcod
       , N'BR'
       , (SELECT DdEmpresa.CdEmpresa FROM DdEmpresa)
       , (SELECT DdEmpresa.CdFilial FROM DdEmpresa)
       , N'Garça'
       , N'SP'
       , N'M'
       , N'3516705'
       , N''
       , GETDATE()
       , 0);


INSERT INTO CEP_Bairro
( chave_bairro
, chave_local
, cdpais
, cdempresa
, cdfilial
, nome
, uf
, updregistro
, stdel)
VALUES ( N'Z000'
       , @newcod
       , N'BR'
       , (SELECT DdEmpresa.CdEmpresa FROM DdEmpresa)
       , (SELECT DdEmpresa.CdFilial FROM DdEmpresa)
       , N'Centro'
       , N'SP'
       , GETDATE()
       , 0);

IF (@cep >= 1)

    BEGIN

        UPDATE CEP
        SET cep             = N'17400000'
          , uf              = N'SP'
          , cdpais          = N'BR'
          , cdempresa       = (SELECT DdEmpresa.CdEmpresa FROM DdEmpresa)
          , cdfilial        = (SELECT DdEmpresa.CdFilial FROM DdEmpresa)
          , chave_local     = @newcod
          , nome_logradouro = N'R. Cel. Joaquim Piza'
          , chave_bairroi   = N''
          , chave_bairrof   = N''
          , complemento_log = N''
          , chave_log       = N'Z000'
          , cidade          = N''
          , updregistro     = GETDATE()
          , stdel           = 0
        WHERE cep = '17400000'

    END

ELSE

    BEGIN

        INSERT INTO CEP ( cep
                        , uf
                        , cdpais
                        , cdempresa
                        , cdfilial
                        , chave_local
                        , nome_logradouro
                        , chave_bairroi
                        , chave_bairrof
                        , complemento_log
                        , chave_log
                        , cidade
                        , updregistro
                        , stdel)
        VALUES ( N'17400000'
               , N'SP'
               , N'BR'
               , (SELECT DdEmpresa.CdEmpresa FROM DdEmpresa)
               , (SELECT DdEmpresa.CdFilial FROM DdEmpresa)
               , @newcod
               , N'R. Cel. Joaquim Piza'
               , N''
               , N''
               , N''
               , N'Z000'
               , N''
               , GETDATE()
               , 0)

    END

