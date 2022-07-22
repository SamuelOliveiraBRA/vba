Function validarCPF(numeroCPF)
On Error Goto TratarErro

'========================================================
' VALIDAÇÃO DE CPF                                      =
'========================================================
' OBJETIVO:         Módulo de validação de CPF          =
' DESENVOLVEDOR:    Samuel Oliveira                     =
' Copyright 2021 Todos os direitos reservados           =
'========================================================

validado = False

'DEIXAR APENAS OS NÚMEROS DO CPF
For posicao = 1 To Len(numeroCPF)
    If Mid(numeroCPF, posicao, 1) Like "[0-9]" Then CPF = CPF & Mid(numeroCPF, posicao, 1)
Next

'SE O CPF não tiver números ou tamanho é menor que 11
If IsNull(IIf(CPF = "", Null, CPF)) Or (Len(CPF) <> 11) Then Exit Function

'lista de cpfs inválidos ou bloqueados
cpfInvalidos = "00000000000,11111111111,22222222222,33333333333,44444444444,55555555555,66666666666,777777777777,88888888888,99999999999)"

For etapa = 1 To 2

    multiplicador = IIf(etapa = 1, 10, 11)  'Decréscimo ou multiplicadores
    TamanhoCPF = IIf(etapa = 1, 9, 10)      'tamanho para o cálculo
    
    For posicao = 1 To TamanhoCPF
        CalculoDigito = CalculoDigito + Mid(CPF, posicao, 1) * multiplicador: multiplicador = multiplicador - 1
    Next posicao
        
    CalculoDigito = IIf(CalculoDigito Mod 11 < 2, 0, 11 - (CalculoDigito Mod 11))
    strCPF = strCPF & CalculoDigito
    CalculoDigito = 0

Next etapa

If CStr(Right(CPF, 2)) = CStr(strCPF) Then validado = True

If InStr(1, cpfInvalidos, CPF) > 0 Then validado = False

validarCPF = validado

SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbExclamation, " Erro " & Err.Number
Resume SairFunction
End Function
