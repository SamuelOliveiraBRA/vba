Function CriarDiretorio(Caminho)
On Error GoTo TratarErro

'========================================================
' CRIAÇÃO DE DIRETÓRIOS DINÂMICOS                       =
'========================================================
' OBJETIVO:         Módul para criar diretórios         =
' DESENVOLVEDOR:    Samuel Oliveira                     =
' CONTATO:          samuel.santos@oprograma.com         =
' WEB SITE          www.oprograma.com                   =
' Copyright 2020 Todos os direitos reservados           =
'========================================================
' OBSERVAÇÕES                                           =
'========================================================

If IsNull(IIf(Caminho = "", Null, Caminho)) Then Exit Function

TamanhoPastas = UBound(Split(Caminho, "\"))

NovoDiretorio = Split(Caminho, "\")(0)
For Etapa = 1 To TamanhoPastas
    If Dir(NovoDiretorio & "\" & Split(Caminho, "\")(Etapa), vbDirectory) = "" Then MkDir NovoDiretorio & "\" & Split(Caminho, "\")(Etapa)
    NovoDiretorio = NovoDiretorio & "\" & Split(Caminho, "\")(Etapa)
Next Etapa

CriarDiretorio = NovoDiretorio

SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbCritical, " Erro " & Err.Number
Resume SairFunction
End Function

