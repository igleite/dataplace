Dim var

var = "Ol� Mundo!"

Copia var

'Copia o c�digo para o CTRL + C
Private Function Copia(ByVal aqui)
    
    Dim WshShell, oExec, oIn
    Set WshShell = CreateObject("WScript.Shell")
    Set oExec = WshShell.Exec("clip")
    Set oIn = oExec.stdIn
    oIn.WriteLine aqui
    oIn.Close

    MsgBox "Copiado para o CTRL + C"

End Function