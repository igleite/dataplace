/*

O comando acima é bem prático e funciona sem problemas. Entretanto, pode ocorrer de entre o tempo de você eliminar as sessões e executar o seu RESTORE DATABASE, por exemplo, outras conexões se conectem no banco. Por este motivo, o método mais recomendável para essa operação é com ALTER DATABASE, colocando o banco de dados no modo SINGLE_USER, onde somente um único usuário pode se conectar por vez:

*/

ALTER DATABASE Testes SET SINGLE_USER WITH ROLLBACK IMMEDIATE
GO


/*

O parâmetro WITH ROLLBACK IMMEDIATE faz com que todas as sessões sejam fechadas sem qualquer aviso e feito o rollback imediatamente. Agora você pode realizar suas manutenções sem possibilidade de qualquer outra conexão interferindo nos seus comandos. Ao final da sua manutenção, lembre-se de voltar o banco para o modo MULTI_USER, para que ele volte a aceitar múltiplas conexões novamente:

*/

ALTER DATABASE Testes SET MULTI_USER
GO