Dim conn
Dim odbc
Dim user
Dim password

odbc        = "sym_TREINAMENTO"
user        = "sa"
password    = "dp"

Set conn = CreateObject("ADODB.Connection")
conn.Open "DSN=" & odbc & ";Persist Security Info=true;Uid=" & user & ";Pwd=" & password & ";"
conn.CommandTimeout = 0
conn.CursorLocation = 3



If conn.State = 1 Then

    Dim ano
    Dim numOrdemProducao
    Dim seqOP
    Dim seqOP2
    
    ano                 =   "2022" 
    numOrdemProducao    =   "17018" 
    seqOP               =   "1" 
    seqOP2              =   "1"

    GerarEtiquetasCustom ano, numOrdemProducao, seqOP, seqOP2
    
End If



Private Function GerarEtiquetasCustom(ByRef ano, ByRef numOrdemProducao, ByRef seqOP, ByRef seqOP2)
    
    Dim strSQL
    Dim orsAux
    Dim rdexecdirect
    Dim zpl
    Dim prn
    Dim produto
    Dim quantidade
    Dim unidade
    Dim lote
    Dim etiqueta

    produto = ""
    quantidade = ""
    unidade = ""
    lote = ""
    etiqueta = 1


    Set orsAux = Createobject("ADODB.recordset")

    ' strSQL = strSQL & vbNewLine & "DECLARE @ChunkLength INT=26;                                                                                                                                                                           "
    strSQL = strSQL & vbNewLine & "SELECT quantidade = IIF(UPPER(LTRIM(RTRIM(ordemproducaoitem.unidade))) = 'KG', CAST(OrdemProducaoItem.QtdEmpenhada AS DECIMAL(19, 8)) * 1000, OrdemProducaoItem.QtdEmpenhada)                            "
    strSQL = strSQL & vbNewLine & "     , unidade    = IIF(UPPER(LTRIM(RTRIM(ordemproducaoitem.unidade))) = 'KG', 'G', LTRIM(RTRIM(ordemproducaoitem.unidade)))                                                                             "
    strSQL = strSQL & vbNewLine & "     , lote       = IIF(LTRIM(RTRIM(OrdemProducaoItemSubUP.SubUPOrigem)) = loteTemporario.nomeLote, NULL, LTRIM(RTRIM(OrdemProducaoItemSubUP.SubUPOrigem)))                                              "
    strSQL = strSQL & vbNewLine & "     , Col1       = descricaoQuebrada.Col1                                                                                                                                                               "
    strSQL = strSQL & vbNewLine & "     , Col2       = descricaoQuebrada.Col2                                                                                                                                                               "
    strSQL = strSQL & vbNewLine & "FROM ordemproducao                                                                                                                                                                                       "
    strSQL = strSQL & vbNewLine & "         LEFT JOIN ordemproducaoitem                                                                                                                                                                     "
    strSQL = strSQL & vbNewLine & "                   ON ordemproducao.ano = ordemproducaoitem.ano                                                                                                                                          "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.numOrdemProducao = ordemproducaoitem.numOrdemProducao                                                                                                           "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.seqOP = ordemproducaoitem.seqOP                                                                                                                                 "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.seqOP2 = ordemproducaoitem.seqOP2                                                                                                                               "
    strSQL = strSQL & vbNewLine & "         LEFT JOIN produto                                                                                                                                                                               "
    strSQL = strSQL & vbNewLine & "                   ON ordemproducaoitem.cdproduto = produto.cdproduto                                                                                                                                    "
    strSQL = strSQL & vbNewLine & "         LEFT JOIN pendenciaop                                                                                                                                                                           "
    strSQL = strSQL & vbNewLine & "                   ON ordemproducao.ano = pendenciaop.ano                                                                                                                                                "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.numOrdemProducao = pendenciaop.numOrdemProducao                                                                                                                 "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.seqOP = pendenciaop.seqOP                                                                                                                                       "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducao.seqOP2 = pendenciaop.seqOP2                                                                                                                                     "
    strSQL = strSQL & vbNewLine & "         LEFT JOIN ordemproducaoitemsubup                                                                                                                                                                "
    strSQL = strSQL & vbNewLine & "                   ON ordemproducaoitemsubup.ano = ordemproducaoitem.ano                                                                                                                                 "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducaoitemsubup.numOrdemProducao = ordemproducaoitem.numOrdemProducao                                                                                                  "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducaoitemsubup.seqOP = ordemproducaoitem.seqOP                                                                                                                        "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducaoitemsubup.seqOP2 = ordemproducaoitem.seqOP2                                                                                                                      "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducaoitemsubup.cdproduto = ordemproducaoitem.cdproduto                                                                                                                "
    strSQL = strSQL & vbNewLine & "                       AND ordemproducaoitemsubup.uporigem = ordemproducaoitem.uporigem                                                                                                                  "
    strSQL = strSQL & vbNewLine & "         OUTER APPLY (                                                                                                                                                                                   "
    strSQL = strSQL & vbNewLine & "    SELECT TOP 1 nomeLote = LTRIM(RTRIM(NumLoteTemporario))                                                                                                                                              "
    strSQL = strSQL & vbNewLine & "    FROM prmsge tmp                                                                                                                                                                                      "
    strSQL = strSQL & vbNewLine & "    WHERE tmp.CdEmpresa = OrdemProducao.CdEmpresa                                                                                                                                                        "
    strSQL = strSQL & vbNewLine & "      AND tmp.CdFilial = OrdemProducao.CdFilial                                                                                                                                                          "
    strSQL = strSQL & vbNewLine & ") loteTemporario                                                                                                                                                                                         "
    
    'DDL da tabela Produto | DsVenda nvarchar(120) | para quebrar a descricao, faço uma pivot retornando a descrição em 4 colunas entao fica | 120 / 4 = 26 (@ChunkLength)
    'Nesse caso aqui vai ser apenas duas colunas
    strSQL = strSQL & vbNewLine & "OUTER APPLY (                                                                                                                                                                                            "
    strSQL = strSQL & vbNewLine & "    SELECT p.*                                                                                                                                                                                           "
    strSQL = strSQL & vbNewLine & "    FROM (                                                                                                                                                                                               "
    strSQL = strSQL & vbNewLine & "             SELECT t.CdProduto                                                                                                                                                                          "
    strSQL = strSQL & vbNewLine & "                  , ColumnName   = CONCAT('Col', A.Nmbr)                                                                                                                                                 "
    strSQL = strSQL & vbNewLine & "                  , Chunk        = SUBSTRING(CONCAT(LTRIM(RTRIM(t.CdProduto)), ' - ', LTRIM(RTRIM(t.DsVenda))), (A.Nmbr - 1) * 26 /*@ChunkLength*/ + 1, 26 /*@ChunkLength*/)                             "
    strSQL = strSQL & vbNewLine & "             FROM Produto t                                                                                                                                                                              "
    strSQL = strSQL & vbNewLine & "                      CROSS APPLY                                                                                                                                                                        "
    strSQL = strSQL & vbNewLine & "                  (                                                                                                                                                                                      "
    strSQL = strSQL & vbNewLine & "                      SELECT TOP ((LEN(CONCAT(LTRIM(RTRIM(t.CdProduto)), ' - ', LTRIM(RTRIM(t.DsVenda)))) + (26 /*@ChunkLength*/ - 1)) / 26 /*@ChunkLength*/) ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) "
    strSQL = strSQL & vbNewLine & "                      FROM master..spt_values                                                                                                                                                            "
    strSQL = strSQL & vbNewLine & "                  ) A(Nmbr)                                                                                                                                                                              "
    strSQL = strSQL & vbNewLine & "             WHERE T.CdProduto = OrdemProducaoItem.CdProduto                                                                                                                                             "
    strSQL = strSQL & vbNewLine & "         ) src                                                                                                                                                                                           "
    strSQL = strSQL & vbNewLine & "             PIVOT                                                                                                                                                                                       "
    strSQL = strSQL & vbNewLine & "             (                                                                                                                                                                                           "
    strSQL = strSQL & vbNewLine & "             MAX(Chunk) FOR ColumnName IN (Col1,Col2 /*add the maximum column count here*/)                                                                                                              "
    strSQL = strSQL & vbNewLine & "             ) p                                                                                                                                                                                         "
    strSQL = strSQL & vbNewLine & ") descricaoQuebrada                                                                                                                                                                                      "
    
    strSQL = strSQL & vbNewLine & "WHERE OrdemProducao.ano                 = '" & ano               & "'"
    strSQL = strSQL & vbNewLine & "  AND OrdemProducao.numOrdemProducao    = '" & numOrdemProducao  & "'"
    strSQL = strSQL & vbNewLine & "  AND OrdemProducao.seqOP               = '" & seqOP             & "'"
    strSQL = strSQL & vbNewLine & "  AND OrdemProducao.seqOP2              = '" & seqOP2            & "'"

    ' Copy strSQL

    orsAux.open strSQL, conn, 3, 3

    If orsAux.recordCount > 0 Then
    
        
        orsAux.MoveFirst

        Do

            quantidade  = orsAux.fields("quantidade").value
            unidade     = orsAux.fields("unidade").value
            lote        = orsAux.fields("lote").value

            col1        = orsAux.fields("Col1").value
            col2        = orsAux.fields("Col2").value

            If etiqueta = 1 Then
            
                'ETIQUETA 1

                zpl = zpl & vbNewLine &"^XA~TA000~JSN^LT0^MNW^MTT^PON^PMN^LH0,0^JMA^PR2,2~SD15^JUS^LRN^CI0^XZ "
                zpl = zpl & vbNewLine &"^XA "
                zpl = zpl & vbNewLine & "^CI28"
                zpl = zpl & vbNewLine &"^MMT "
                zpl = zpl & vbNewLine &"^PW831 "
                zpl = zpl & vbNewLine &"^LL0240 "
                zpl = zpl & vbNewLine &"^LS0 "
                zpl = zpl & vbNewLine & "^CI28"


                
                If Not IsNull(col1) And IsNull(col2) Then
                    '********************************************
                    'Modelo01 1 colunas para descrição do produto
                    '********************************************

                    zpl = zpl & vbNewLine &"^FT794,50^A0I,23,24^FH\^FDOP:" & ano & "." & numOrdemProducao & "." & seqOP &  "." & seqOP2 & "^FS "

                    If IsNull(lote) Then

                        zpl = zpl & vbNewLine &"^FT794,94^A0I,23,24^FH\^FDLote:" & RepeatString(10, "_")  & "^FS"
                    
                    Else

                        zpl = zpl & vbNewLine &"^FT794,94^A0I,23,24^FH\^FDLote:" & lote  & "^FS"

                    End If
                    
                    zpl = zpl & vbNewLine &"^FT794,137^A0I,23,24^FH\^FDQtde:"   & FormatNumber(quantidade, 2) & RepeatString(5, "_") & "^FS"
                    zpl = zpl & vbNewLine &"^FT794,184^A0I,23,24^FH\^FDCd: "    & col1                                               & "^FS"


                ElseIf Not IsNull(col1) And Not IsNull(col2) Then
                    '********************************************
                    'Modelo02 2 colunas para descrição do produto
                    '********************************************

                    zpl = zpl & vbNewLine &"^FT794,32^A0I,23,24^FH\^FDOP:" & ano & "." & numOrdemProducao & "." & seqOP &  "." & seqOP2 & "^FS "

                    If IsNull(lote) Then

                        zpl = zpl & vbNewLine &"^FT794,73^A0I,23,24^FH\^FDLote:" & RepeatString(10, "_")  & "^FS"

                    Else

                        zpl = zpl & vbNewLine &"^FT794,73^A0I,23,24^FH\^FDLote:" & lote  & "^FS"

                    End If

                    
                    zpl = zpl & vbNewLine &"^FT794,114^A0I,23,24^FH\^FDQtde:"   & FormatNumber(quantidade, 2) & RepeatString(5, "_") & "^FS"
                    zpl = zpl & vbNewLine &"^FT794,193^A0I,23,24^FH\^FDCd: "    & col1 & "^FS"
                    zpl = zpl & vbNewLine &"^FT794,152^A0I,23,24^FH\^FD"        & col2 & "^FS"


                End If


                'troca para a etiqueta 2
                etiqueta = 2

            Else

                'ETIQUETA 2

                If Not IsNull(col1) And IsNull(col2) Then
                    '********************************************
                    'Modelo01 1 colunas para descrição do produto
                    '********************************************

                    zpl = zpl & vbNewLine &"^FT370,50^A0I,23,24^FH\^FDOP:" & ano & "." & numOrdemProducao & "." & seqOP &  "." & seqOP2 & "^FS "

                    If IsNull(lote) Then

                        zpl = zpl & vbNewLine &"^FT370,94^A0I,23,24^FH\^FDLote:" & RepeatString(10, "_")  & "^FS"

                    Else

                        zpl = zpl & vbNewLine &"^FT370,94^A0I,23,24^FH\^FDLote:" & lote  & "^FS"

                    End If


                    zpl = zpl & vbNewLine &"^FT370,137^A0I,23,24^FH\^FDQtde:"   & FormatNumber(quantidade, 2) & RepeatString(5, "_") & "^FS"
                    zpl = zpl & vbNewLine &"^FT370,184^A0I,23,24^FH\^FDCd:"     & col1                                               & "^FS"


                ElseIf Not IsNull(col1) And Not IsNull(col2) Then
                    '********************************************
                    'Modelo02 2 colunas para descrição do produto
                    '********************************************

                    zpl = zpl & vbNewLine &"^FT370,32^A0I,23,24^FH\^FDOP:" & ano & "." & numOrdemProducao & "." & seqOP &  "." & seqOP2 & "^FS "

                    If IsNull(lote) Then

                        zpl = zpl & vbNewLine &"^FT370,73^A0I,23,24^FH\^FDLote:" & RepeatString(10, "_")  & "^FS"
                    
                    Else
                    
                        zpl = zpl & vbNewLine &"^FT370,73^A0I,23,24^FH\^FDLote:" & lote  & "^FS"

                    End If

                    zpl = zpl & vbNewLine &"^FT370,114^A0I,23,24^FH\^FDQtde:"   & FormatNumber(quantidade, 2) & RepeatString(5, "_") & "^FS"
                    zpl = zpl & vbNewLine &"^FT370,193^A0I,23,24^FH\^FDCd:"     & col1 & "^FS "
                    zpl = zpl & vbNewLine &"^FT370,152^A0I,23,24^FH\^FD"        & col2 & "^FS"

                End If

                'fecha o rodapé do codigo zpl
                zpl = zpl & vbNewLine &"^PQ1,0,1,Y^XZ"

                'troca para a etiqueta 1
                etiqueta = 1

            End If

            orsAux.moveNext

        Loop Until orsAux.EOF


        'fecha o rodapé do código zpl
        If etiqueta = 2 Then

            zpl = zpl & vbNewLine &"^PQ1,0,1,Y^XZ"

        End If


        prn = zpl

    End if

    Copy prn
    
    'prm.StrReturn = prn    

End Function


'Repete conforme a quantidade solicitada
Private Function RepeatString(ByVal number, ByVal text)

    Redim buffer(number)
    RepeatString = Join( buffer, text )

End Function


'Copia o código zpl para o CTRL + C
Private Function Copy(ByVal aqui)
    
    Dim WshShell, oExec, oIn
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec("clip")
    Set oIn = oExec.stdIn
    oIn.WriteLine aqui
    oIn.Close

    MsgBox "zpl Copiado"

    OpenzplEmulator

End Function


'Abre site para testar o código zpl Online
Private Function OpenzplEmulator()
    
    Dim variable
    Dim urlForOpen

    urlForOpen = "https://zplprinter.azurewebsites.net/"
    Set variable = CreateObject("WScript.Shell")
    variable.Run urlForOpen

End Function