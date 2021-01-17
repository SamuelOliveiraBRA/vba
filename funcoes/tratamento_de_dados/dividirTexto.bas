Function DividirTexto(Texto, Divisor, Posicao)
On Error GoTo TratarErro

If Not IsNull(IIf(Texto = "", Null, Texto)) Then DividirTexto = Split(Texto, Divisor)(Posicao)

DividirTexto = Split(Texto, Divisor)(Posicao)

SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbCritical, " Erro " & Err.Number
Resume SairFunction
End Function
