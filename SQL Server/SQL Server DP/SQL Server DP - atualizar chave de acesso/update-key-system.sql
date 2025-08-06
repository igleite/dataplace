USE SPT_IGO_DbPassoFundo
GO

SELECT CdNaDp, Chave, *
FROM DdEmpresa

UPDATE DdEmpresa
SET Chave = 'nova chave aqui'
OUTPUT deleted.Chave AS 'ChaveAntiga',
    inserted.Chave AS 'ChaveNova'
