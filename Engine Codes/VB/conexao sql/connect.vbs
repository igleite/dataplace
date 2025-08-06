Dim conn
Dim odbc
Dim user
Dim password

odbc        = "sym_treinamento"
user        = "sa"
password    = "dp"

Set conn = CreateObject("ADODB.Connection")
conn.Open "DSN=" & odbc & ";Persist Security Info=true;Uid=" & user & ";Pwd=" & password & ";"
conn.CommandTimeout = 0
conn.CursorLocation = 3



If conn.State = 1 Then

    MsgBox "Conex�o OK!"
    
End If