' ************************************************************
Dim odbc
Dim user
Dim password
Dim ocnADO
Dim ncm
Dim dt_inicio, dt_fim
Dim excUpd
excUpd = True
' ************************************************************




' INFORMAR DSN, USUARIO E SENHA, AQUI
odbc        = "sym_treinamento"
user        = "sa"
password    = "dp"




' ADICIONAR DATAS, AQUI. UTILIZADAS PARA ATUALIZAÇÃO DA TABELA PendenciaVendaItem e PedidoItem
dt_inicio   = "2022-04-27"
dt_fim      = "2022-04-28"



' ADICIONAR NCM E ALQ. IPI, AQUI. DA MESMA FORMA QUE ESTÁ ABAIXO NO EXEMPLO
' (NCM||alq. Nova|Alq. Antiga)
ncm =   "(01012100||3.75|5)"    _
&       "(73072200||3.75|5)"    _






' ************************************************************
' abre a conexão
Set ocnADO = CreateObject("ADODB.Connection")
ocnADO.Open "DSN=" & odbc & ";Persist Security Info=true;Uid=" & user & ";Pwd=" & password & ";"
ocnADO.CommandTimeout = 0
ocnADO.CursorLocation = 3


' se a conexão estiver ok, chama as funções
If ocnADO.State = 1 Then


    ' criação das tabelas de auxilio
    createTrocaAlqIpiPeriodo
    createTrocaAlqIpi

    ' inserção dos dados na tabela de auxilio
    insertTrocaAlqIpiPeriodo    dt_inicio, dt_fim
    insertTrocaAlqIpi           ncm


    If MsgBox ("Deseja atualizar a Alq. de IPI da tabela (ExcecaoIcms)?", vbQuestion + vbYesNo + vbSystemModal, "Selecione uma opção!") = vbYes then

        backupExcecaoICMS
        updateExcecaoICMS
        
    End If


    If MsgBox ("Deseja atualizar a Alq. de IPI da tabela (PendenciaVendaItem)?", vbQuestion + vbYesNo + vbSystemModal, "Selecione uma opção!") = vbYes then

        backupPendenciaVendaItem
        updatePendenciaVendaItem
        
    End If


    If MsgBox ("Deseja atualizar a Alq. de IPI da tabela (PedidoItem)?", vbQuestion + vbYesNo + vbSystemModal, "Selecione uma opção!") = vbYes then

        backupPedidoItem
        updatePedidoItem
        
    End If


    If MsgBox ("Deseja atualizar a Alq. de IPI da tabela (Produto)?", vbQuestion + vbYesNo + vbSystemModal, "Selecione uma opção!") = vbYes then

        backupProduto
        updateProduto
        
    End If


    If MsgBox ("Deseja atualizar a Alq. de IPI da tabela (CENCM)?", vbQuestion + vbYesNo + vbSystemModal, "Selecione uma opção!") = vbYes then

        backupCENCM
        updateCENCM
        
    End If
    

End If


' cria a tabela de auxilio | datas
Private Function createTrocaAlqIpiPeriodo()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS dp_trocaAlqIpiPeriodo  "
        strSQL = strSQL & vbNewLine & "CREATE TABLE dp_trocaAlqIpiPeriodo          "
        strSQL = strSQL & vbNewLine & "(                                           "
        strSQL = strSQL & vbNewLine & "    id        INT IDENTITY (1,1),           "
        strSQL = strSQL & vbNewLine & "    dt_inicio DATE NOT NULL,                "
        strSQL = strSQL & vbNewLine & "    dt_fim    DATE NOT NULL                 "
        strSQL = strSQL & vbNewLine & ")                                           "

        ' Copy strSQL
        ocnADO.Execute strSQL, rdexecdirect




        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' cria a tabela de auxilio | ncm e Alq. Ipi
Private Function createTrocaAlqIpi()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS dp_trocaAlqIpi                  "
        strSQL = strSQL & vbNewLine & "CREATE TABLE dp_trocaAlqIpi                          "
        strSQL = strSQL & vbNewLine & "(                                                    "
        strSQL = strSQL & vbNewLine & "    id           INT IDENTITY (1,1),                 "
        strSQL = strSQL & vbNewLine & "    ncm          VARCHAR(8)    NOT NULL UNIQUE,      "
        strSQL = strSQL & vbNewLine & "    alq_ipi      DECIMAL(7, 2) NOT NULL,             "
        strSQL = strSQL & vbNewLine & "    alq_ipi_old  DECIMAL(7, 2) NOT NULL              "
        strSQL = strSQL & vbNewLine & ")                                                    "

        ' Copy strSQL
        ocnADO.Execute strSQL, rdexecdirect




        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' insere as datas na tabela de auxilio
Private Function insertTrocaAlqIpiPeriodo(ByVal dt_inicio, ByVal dt_fim)
    On Error Resume Next


        Dim strSQL, rdexecdirect, ncm2

        strSQL = strSQL & vbNewLine & "INSERT INTO dp_trocaAlqIpiPeriodo (dt_inicio, dt_fim) VALUES ('" & dt_inicio & "','" & dt_fim & "')"

        ' Copy strSQL
        ocnADO.Execute strSQL, rdexecdirect




        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' insere ncm na tabela de auxilio
Private Function insertTrocaAlqIpi(ByVal ncm)
    On Error Resume Next

        Dim strSQL, rdexecdirect

        ncm = Replace(ncm, "|", ",") 
        ncm = Replace(ncm, ",,", "', ")
        ncm = Replace(ncm, "(", "('")
        ncm = Replace(ncm, ")(", "),(")
        

        strSQL = "INSERT INTO dp_trocaAlqIpi (ncm, alq_ipi, alq_ipi_old) VALUES" & ncm

        ' Copy strSQL
        ocnADO.Execute strSQL, rdexecdirect




        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' ************************************************************
' BACKUPS
' ************************************************************

' faz backup da tabela tabela ExcecaoICMS, com base na tabela de auxilio
Private Function backupExcecaoICMS() 
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS ExcecaoICMS_dp_trocaAlqIpi                      "
        strSQL = strSQL & vbNewLine & "SELECT ExcecaoICMS.*                                                 "
        strSQL = strSQL & vbNewLine & "INTO ExcecaoICMS_dp_trocaAlqIpi                                      "
        strSQL = strSQL & vbNewLine & "FROM Produto                                                         "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                "
        strSQL = strSQL & vbNewLine & "                    ON Produto.ncm = tAI.ncm                         "
        strSQL = strSQL & vbNewLine & "        INNER JOIN ExcecaoICMS                                       "
        strSQL = strSQL & vbNewLine & "                    ON Produto.CdProduto = ExcecaoICMS.CdProduto     "
        strSQL = strSQL & vbNewLine & "                        AND tAI.alq_ipi_old = ExcecaoICMS.alqipi     "

        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If

        MsgBox "Backup da tabela (ExcecaoICMS) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' faz backup da tabela tabela PendenciaVendaItem, com base na tabela de auxilio
Private Function backupPendenciaVendaItem()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS pendenciavendaitem_dp_trocaAlqIpi                                                           "
        strSQL = strSQL & vbNewLine & "SELECT PVI.*                                                                                                     "
        strSQL = strSQL & vbNewLine & "INTO pendenciavendaitem_dp_trocaAlqIpi                                                                           "
        strSQL = strSQL & vbNewLine & "FROM pendenciavenda1 PV                                                                                          "
        strSQL = strSQL & vbNewLine & "        INNER JOIN pendenciavendaitem PVI                                                                        "
        strSQL = strSQL & vbNewLine & "                    ON PV.cdempresa = PVI.cdempresa                                                              "
        strSQL = strSQL & vbNewLine & "                        AND PV.cdfilial = PVI.cdfilial                                                           "
        strSQL = strSQL & vbNewLine & "                        AND PV.nummovimento = PVI.nummovimento                                                   "
        strSQL = strSQL & vbNewLine & "        INNER JOIN Produto P                                                                                     "
        strSQL = strSQL & vbNewLine & "                    ON PVI.cdempresa = P.CdEmpresa                                                               "
        strSQL = strSQL & vbNewLine & "                        AND PVI.cdfilial = P.CdFilial                                                            "
        strSQL = strSQL & vbNewLine & "                        AND PVI.cdproduto = P.CdProduto                                                          "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                            "
        strSQL = strSQL & vbNewLine & "                    ON P.ncm = tAI.ncm                                                                           "
        strSQL = strSQL & vbNewLine & "                        AND PVI.alqipi = tAI.alq_ipi_old                                                         "
        strSQL = strSQL & vbNewLine & "        OUTER APPLY (                                                                                            "
        strSQL = strSQL & vbNewLine & "    SELECT TOP 1 dt_inicio                                                                                       "
        strSQL = strSQL & vbNewLine & "           , dt_fim                                                                                              "
        strSQL = strSQL & vbNewLine & "    FROM dp_trocaAlqIpiPeriodo                                                                                   "
        strSQL = strSQL & vbNewLine & "    ORDER BY id DESC                                                                                             "
        strSQL = strSQL & vbNewLine & ") data                                                                                                           "
        strSQL = strSQL & vbNewLine & "WHERE PVI.qtdprocessada <> PVI.qtdsolicitada                                                                     "
        strSQL = strSQL & vbNewLine & "AND PV.dtinicio BETWEEN CONCAT(data.dt_inicio, ' ', '00:00:00.000') AND CONCAT(data.dt_fim, ' ', '23:59:00.000') "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Backup da tabela (PendenciaVendaItem) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' faz backup da tabela tabela PedidoItem, com base na tabela de auxilio
Private Function backupPedidoItem()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS PedidoItem_dp_trocaAlqIpi                                                                   "
        strSQL = strSQL & vbNewLine & "SELECT PII.*                                                                                                      "
        strSQL = strSQL & vbNewLine & "INTO PedidoItem_dp_trocaAlqIpi                                                                                   "
        strSQL = strSQL & vbNewLine & "FROM Pedido P                                                                                                    "
        strSQL = strSQL & vbNewLine & "        INNER JOIN PedidoItem PII                                                                                 "
        strSQL = strSQL & vbNewLine & "                    ON P.cdempresa = PII.cdempresa                                                                "
        strSQL = strSQL & vbNewLine & "                        AND P.cdfilial = PII.cdfilial                                                             "
        strSQL = strSQL & vbNewLine & "                        AND P.NumPedido = PII.numpedido                                                           "
        strSQL = strSQL & vbNewLine & "        INNER JOIN Produto PP                                                                                    "
        strSQL = strSQL & vbNewLine & "                    ON PII.cdempresa = PP.CdEmpresa                                                               "
        strSQL = strSQL & vbNewLine & "                        AND PII.cdfilial = PP.CdFilial                                                            "
        strSQL = strSQL & vbNewLine & "                        AND PII.cdproduto = PP.CdProduto                                                          "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                            "
        strSQL = strSQL & vbNewLine & "                    ON PP.ncm = tAI.ncm                                                                          "
        strSQL = strSQL & vbNewLine & "                        AND PII.alqipi = tAI.alq_ipi_old                                                          "
        strSQL = strSQL & vbNewLine & "        OUTER APPLY (                                                                                            "
        strSQL = strSQL & vbNewLine & "    SELECT TOP 1 dt_inicio                                                                                       "
        strSQL = strSQL & vbNewLine & "            , dt_fim                                                                                             "
        strSQL = strSQL & vbNewLine & "    FROM dp_trocaAlqIpiPeriodo                                                                                   "
        strSQL = strSQL & vbNewLine & "    ORDER BY id DESC                                                                                             "
        strSQL = strSQL & vbNewLine & ") data                                                                                                           "
        strSQL = strSQL & vbNewLine & "WHERE P.StPedido = 'A'                                                                                           "
        strSQL = strSQL & vbNewLine & "AND P.DtPedido BETWEEN CONCAT(data.dt_inicio, ' ', '00:00:00.000') AND CONCAT(data.dt_fim, ' ', '23:59:00.000')  "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Backup da tabela (PedidoItem) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' faz backup da tabela tabela Produto, com base na tabela de auxilio
Private Function backupProduto()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS Produto_dp_trocaAlqIpi                                                                      "
        strSQL = strSQL & vbNewLine & "SELECT PP.*                                                                                                      "
        strSQL = strSQL & vbNewLine & "INTO Produto_dp_trocaAlqIpi                                                                                      "
        strSQL = strSQL & vbNewLine & "FROM Produto PP                                                                                                  "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                            "
        strSQL = strSQL & vbNewLine & "                    ON PP.ncm = tAI.ncm                                                                          "
        strSQL = strSQL & vbNewLine & "                        AND PP.alqipi = tAI.alq_ipi_old                                                         "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Backup da tabela (Produto) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' faz backup da tabela tabela CENCM, com base na tabela de auxilio
Private Function backupCENCM()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "DROP TABLE IF EXISTS CENCM_dp_trocaAlqIpi                                                                        "
        strSQL = strSQL & vbNewLine & "SELECT NCM.*                                                                                                     "
        strSQL = strSQL & vbNewLine & "INTO CENCM_dp_trocaAlqIpi                                                                                        "
        strSQL = strSQL & vbNewLine & "FROM CENCM NCM                                                                                                   "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                            "
        strSQL = strSQL & vbNewLine & "                    ON NCM.ncm = tAI.ncm                                                                         "
        strSQL = strSQL & vbNewLine & "                        AND NCM.alqipi = tAI.alq_ipi_old                                                         "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Backup da tabela (CENCM) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' ************************************************************
' END BACKUPS
' ************************************************************


' ************************************************************
' UPDATES
' ************************************************************


' atualiza a alq. ipi da tabela ExcecaoICMS, com base na tabela de auxilio
Private Function updateExcecaoICMS() 
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "UPDATE ExcecaoICMS WITH (ROWLOCK)                                    "
        strSQL = strSQL & vbNewLine & "SET ExcecaoICMS.alqipi = tAI.alq_ipi                                 "
        strSQL = strSQL & vbNewLine & "FROM Produto                                                         "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                "
        strSQL = strSQL & vbNewLine & "                    ON Produto.ncm = tAI.ncm                         "
        strSQL = strSQL & vbNewLine & "        INNER JOIN ExcecaoICMS                                       "
        strSQL = strSQL & vbNewLine & "                    ON Produto.CdProduto = ExcecaoICMS.CdProduto     "
        strSQL = strSQL & vbNewLine & "                        AND tAI.alq_ipi_old = ExcecaoICMS.alqipi     "

        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Atualização de IPI da tabela (ExcecaoICMS) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' atualiza a alq. ipi da tabela PendenciaVendaItem, com base na tabela de auxilio
Private Function updatePendenciaVendaItem()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "UPDATE pendenciavendaitem WITH (ROWLOCK)                                                                             "
        strSQL = strSQL & vbNewLine & "SET alqipi = tAI.alq_ipi                                                                                             "
        strSQL = strSQL & vbNewLine & "FROM pendenciavenda1 PV                                                                                              "
        strSQL = strSQL & vbNewLine & "        INNER JOIN pendenciavendaitem PVI                                                                            "
        strSQL = strSQL & vbNewLine & "                    ON PV.cdempresa = PVI.cdempresa                                                                  "
        strSQL = strSQL & vbNewLine & "                        AND PV.cdfilial = PVI.cdfilial                                                               "
        strSQL = strSQL & vbNewLine & "                        AND PV.nummovimento = PVI.nummovimento                                                       "
        strSQL = strSQL & vbNewLine & "        INNER JOIN Produto P                                                                                         "
        strSQL = strSQL & vbNewLine & "                    ON PVI.cdempresa = P.CdEmpresa                                                                   "
        strSQL = strSQL & vbNewLine & "                        AND PVI.cdfilial = P.CdFilial                                                                "
        strSQL = strSQL & vbNewLine & "                        AND PVI.cdproduto = P.CdProduto                                                              "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                                "
        strSQL = strSQL & vbNewLine & "                    ON P.ncm = tAI.ncm                                                                               "
        strSQL = strSQL & vbNewLine & "                        AND PVI.alqipi = tAI.alq_ipi_old                                                             "
        strSQL = strSQL & vbNewLine & "        OUTER APPLY (                                                                                                "
        strSQL = strSQL & vbNewLine & "    SELECT TOP 1 dt_inicio                                                                                           "
        strSQL = strSQL & vbNewLine & "           , dt_fim                                                                                                  "
        strSQL = strSQL & vbNewLine & "    FROM dp_trocaAlqIpiPeriodo                                                                                       "
        strSQL = strSQL & vbNewLine & "    ORDER BY id DESC                                                                                                 "
        strSQL = strSQL & vbNewLine & ") data                                                                                                               "
        strSQL = strSQL & vbNewLine & "WHERE PVI.qtdprocessada <> PVI.qtdsolicitada                                                                         "
        strSQL = strSQL & vbNewLine & "  AND PV.stpendencia IN ('A', 'F')                                                                                   "
        strSQL = strSQL & vbNewLine & "  AND PV.dtinicio BETWEEN CONCAT(data.dt_inicio, ' ', '00:00:00.000') AND CONCAT(data.dt_fim, ' ', '23:59:00.000')   "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Atualização de IPI da tabela (PendenciaVendaItem) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' atualiza a alq. ipi da tabela PedidoItem, com base na tabela de auxilio
Private Function updatePedidoItem()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "UPDATE PedidoItem WITH (ROWLOCK)                                                                                     "
        strSQL = strSQL & vbNewLine & "SET alqipi = tAI.alq_ipi                                                                                             "
        strSQL = strSQL & vbNewLine & "FROM Pedido P                                                                                                        "
        strSQL = strSQL & vbNewLine & "        INNER JOIN PedidoItem PII                                                                                    "
        strSQL = strSQL & vbNewLine & "                    ON P.cdempresa = PII.cdempresa                                                                   "
        strSQL = strSQL & vbNewLine & "                        AND P.cdfilial = PII.cdfilial                                                                "
        strSQL = strSQL & vbNewLine & "                        AND P.NumPedido = PII.numpedido                                                              "
        strSQL = strSQL & vbNewLine & "        INNER JOIN Produto PP                                                                                        "
        strSQL = strSQL & vbNewLine & "                    ON PII.cdempresa = PP.CdEmpresa                                                                  "
        strSQL = strSQL & vbNewLine & "                        AND PII.cdfilial = PP.CdFilial                                                               "
        strSQL = strSQL & vbNewLine & "                        AND PII.cdproduto = PP.CdProduto                                                             "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                                "
        strSQL = strSQL & vbNewLine & "                    ON PP.ncm = tAI.ncm                                                                              "
        strSQL = strSQL & vbNewLine & "                        AND PII.alqipi = tAI.alq_ipi_old                                                             "
        strSQL = strSQL & vbNewLine & "        OUTER APPLY (                                                                                                "
        strSQL = strSQL & vbNewLine & "    SELECT TOP 1 dt_inicio                                                                                           "
        strSQL = strSQL & vbNewLine & "            , dt_fim                                                                                                 "
        strSQL = strSQL & vbNewLine & "    FROM dp_trocaAlqIpiPeriodo                                                                                       "
        strSQL = strSQL & vbNewLine & "    ORDER BY id DESC                                                                                                 "
        strSQL = strSQL & vbNewLine & ") data                                                                                                               "
        strSQL = strSQL & vbNewLine & "WHERE P.StPedido = 'A'                                                                                               "
        strSQL = strSQL & vbNewLine & "AND P.DtPedido BETWEEN CONCAT(data.dt_inicio, ' ', '00:00:00.000') AND CONCAT(data.dt_fim, ' ', '23:59:00.000')      "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Atualização de IPI da tabela (PedidoItem) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' atualiza a alq. ipi da tabela Produto, com base na tabela de auxilio
Private Function updateProduto()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "UPDATE Produto WITH (ROWLOCK)                                                                                        "
        strSQL = strSQL & vbNewLine & "SET AlqIPI = tAI.alq_ipi                                                                                             "
        strSQL = strSQL & vbNewLine & "FROM Produto PP                                                                                                      "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                                "
        strSQL = strSQL & vbNewLine & "                    ON PP.ncm = tAI.ncm                                                                              "
        strSQL = strSQL & vbNewLine & "                        AND PP.alqipi = tAI.alq_ipi_old                                                              "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Atualização de IPI da tabela (Produto) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' atualiza a alq. ipi da tabela CENCM, com base na tabela de auxilio
Private Function updateCENCM()
    On Error Resume Next

        Dim strSQL, rdexecdirect

        strSQL = strSQL & vbNewLine & "UPDATE CENCM WITH (ROWLOCK)                                                                                           "
        strSQL = strSQL & vbNewLine & "SET AlqIPI = tAI.alq_ipi                                                                                              "
        strSQL = strSQL & vbNewLine & "FROM CENCM NCM                                                                                                        "
        strSQL = strSQL & vbNewLine & "        INNER JOIN dp_trocaAlqIpi tAI                                                                                 "
        strSQL = strSQL & vbNewLine & "                    ON NCM.ncm = tAI.ncm                                                                              "
        strSQL = strSQL & vbNewLine & "                        AND NCM.alqipi = tAI.alq_ipi_old                                                              "


        ' Copy strSQL
        If excUpd Then
            ocnADO.Execute strSQL, rdexecdirect
        End If


        MsgBox "Atualização de IPI da tabela (CENCM) concluído com sucesso!", vbinformation + vbSystemModal, "Informativo"


        ' RETORNA MENSAGEM DE ERRO
        If Err.Number <> 0 Then
            MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
            Exit Function
        End If

    On Error Goto 0
End Function


' ************************************************************
' END UPDATES
' ************************************************************


' copia string para o ctrl + c
Private Function Copy(ByVal aqui)
    
    Dim WshShell, oExec, oIn
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec("clip")
    Set oIn = oExec.stdIn
    oIn.WriteLine aqui
    oIn.Close

    MsgBox "Copiado para o CTRL + C"

End Function