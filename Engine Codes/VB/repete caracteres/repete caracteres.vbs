Dim var

'ir� repetir a letra o  10 vezes
var = "Ol� Mundo" & RepeatString(10, "o")

MsgBox var  ' output: Ol� Mundooooooooooo


'Repete conforme a quantidade solicitada
Private Function RepeatString(ByVal number, ByVal text)

    Redim buffer(number)
    RepeatString = Join( buffer, text )

End Function