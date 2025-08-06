On Error Resume Next

    Dim fso
    Dim gf 
    Dim str
    Dim linhas

    Dim dados
    Dim dtCompensacao
    Dim historico
    Dim docto
    Dim bco
    Dim ag
    Dim conta
    Dim cred
    Dim deb

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set gf = fso.getFile("C:\Temp\_csv\CSV\Bradesco2.CSV")
    Set str = gf.OpenAsTextStream(1, -2)
    Set linhas = CreateObject("System.Collections.ArrayList")



    Do While Not str.AtEndOfStream
        linhas.Add str.ReadLine
    Loop
    str.close


    For i = 1 To linhas.Count


        If i > 1 And i < linhas.Count - 1 Then

            dados = Split(linhas(i), ";")

            dtCompensacao = dados(0)
            historico = dados(1)
            docto = dados(2)
            bco = dados(3)
            ag = dados(4)
            conta = dados(5)
            cred = dados(6)
            deb = dados(7)


            
            MsgBox "Data de Compensação: " & dtCompensacao & _ 
                vbNewLine & "Historico: " & historico & _ 
                vbNewLine & "Documento: " & docto & _ 
                vbNewLine & "Banco: " & bco & _ 
                vbNewLine & "Agencia: " & ag & _
                vbNewLine & "Conta: " & conta & _ 
                vbNewLine & "Credito: " & cre & _ 
                vbNewLine & "Debito: " & deb

        End If


    Next
    

    If Err.Number <> 0 Then
        MsgBox Err.Number &" - "& Err.Description, vbCritical + vbSystemModal, "Informativo"
    End If


On Error Goto 0