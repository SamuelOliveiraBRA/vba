Function validarCNPJ(numeroCNPJ)
On Error Goto TratarErro

'========================================================
' VALIDAÇÃO DE CNPJ                                     =
'========================================================
' OBJETIVO:         Módulo de validação de CPF          =
' DESENVOLVEDOR:    Samuel Oliveira                     =
' CONTATO:          oprograma@oprograma.com             =
' WEB SITE          www.oprograma.com                   =
' Copyright 2014 Todos os direitos reservados           =
'========================================================

validado = False

'DEIXAR APENAS OS NÚMEROS DO CNPJ
For posicao = 1 To Len(numeroCNPJ)
    If Mid(numeroCNPJ, posicao, 1) Like "[0-9]" Then CNPJ = CNPJ & Mid(numeroCNPJ, posicao, 1)
Next

'SE O CNPJ FOR INVÁLIDO OU NÃO TIVER OS 14 CARACTERES
If IsNull(IIf(CNPJ = "", Null, CNPJ)) Or (Len(CNPJ) <> 14) Then Exit Function

cnpjInvalidos = "00000000000000,11111111111111,22222222222222,33333333333333,44444444444444,55555555555555,66666666666666,77777777777777,88888888888888,99999999999999)"

For etapa = 1 To 2

    multiplicador = IIf(etapa = 1, "543298765432", "6543298765432") 'matriz de multiplicação
    TamanhoCNPJ = IIf(etapa = 1, 12, 13)                            'tamanho para o cálculo do dígito
    
    For posicao = 1 To TamanhoCNPJ
        If Mid(CNPJ, posicao, 1) Like "[0-9]" Then
        
            CalculoDigito = CalculoDigito + Mid(CNPJ, posicao, 1) * Mid(multiplicador, posicao, 1)
        
        End If
    Next posicao
    CalculoDigito = IIf(CalculoDigito Mod 11 < 2, 0, 11 - (CalculoDigito Mod 11))
    strCNPJ = strCNPJ & CalculoDigito
    CalculoDigito = 0
    
Next etapa

If CStr(Right(CNPJ, 2)) = CStr(strCNPJ) Then validado = True

If InStr(1, cnpjInvalidos, CNPJ) > 0 Then validado = False

validarCNPJ = validado
          
SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbExclamation, " Erro " & Err.Number
Resume SairFunction
End Function
