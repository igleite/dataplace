BEGIN TRAN
    
    UPDATE Usuario
    SET SenhaAutenticacao = ''
    OUTPUT deleted.SenhaAutenticacao AS 'SenhaOld', inserted.SenhaAutenticacao AS 'SenhaAtualizada'
    WHERE username = 'sym_' + 'edson'

-- ROLLBACK TRAN 
-- COMMIT TRAN