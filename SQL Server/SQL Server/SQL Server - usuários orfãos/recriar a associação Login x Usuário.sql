USE [Testes]
EXEC sp_change_users_login 'Auto_Fix', 'Usuario_Orfao' -- Isso irá associar o Login 'Usuario_Orfao' ao usuário 'Usuario_Orfao'
GO