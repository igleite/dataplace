DECLARE @user NVARCHAR(50)

SET @user = 'sym_' + 'adolfo'

SELECT 'Email'             = LTRIM(RTRIM(email))
     , 'SenhaEmail'        = LTRIM(RTRIM(SenhaAutenticacao))
     , 'ExigeAutenticacao' = LTRIM(RTRIM(StExigeAutenticacao))
     , 'Username'          = LTRIM(RTRIM(username))
     , 'UsernameLogin'     = LTRIM(RTRIM(REPLACE(username, 'sym_', '')))
     , 'Nome'              = LTRIM(RTRIM(nome))
     , 'SenhaSistema'      = LTRIM(RTRIM(senha))
FROM Usuario (NOLOCK)
WHERE username LIKE '%' + @user + '%'

SELECT 'nome'                = LTRIM(RTRIM(nome))
     , 'emailcomprador'      = LTRIM(RTRIM(emailcomprador))
     , 'stexigeautenticacao' = LTRIM(RTRIM(stexigeautenticacao))
     , 'SenhaAutenticacao'   = LTRIM(RTRIM(SenhaAutenticacao))
FROM Comprador
