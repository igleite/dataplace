Dim conn
Dim odbc
Dim user
Dim password
Dim orsAux
Dim strSQL
Dim razao

odbc        = "sym_treinamento"
user        = "sa"
password    = "dp"

Set conn = CreateObject("ADODB.Connection")
conn.Open "DSN=" & odbc & ";Persist Security Info=true;Uid=" & user & ";Pwd=" & password & ";"
conn.CommandTimeout = 0
conn.CursorLocation = 3



If conn.State = 1 Then

    Set orsAux = CreateObject("ADODB.Recordset")

    strSQL = strSQL & vbNewLine & "SELECT TOP 2 Razao    = LTRIM(RTRIM(Razao))"
    strSQL = strSQL & vbNewLine & "           , Fantasia = LTRIM(RTRIM(Fantasia))"
    strSQL = strSQL & vbNewLine & "FROM Empresa"
    
    orsAux.Open strSQL, conn, 3, 3
    


    If orsAux.RecordCount > 0 Then

        orsAux.MoveFirst

        Do

            razao  = orsAux.fields("Razao").value
            fantasia  = orsAux.fields("Fantasia").value

            MsgBox "Razão Social: " & razao & vbNewLine & "Fantasia: " & fantasia

            orsAux.moveNext
        Loop Until orsAux.EOF

    End If
    


End If