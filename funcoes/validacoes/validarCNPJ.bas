Function validarCNPJ(numeroCNPJ As String) As Boolean
    On Error GoTo TratarErro
    
    '========================================================
    ' VALIDAÇÃO DE CNPJ (NUMÉRICO OU ALFANUMÉRICO)          =
    '========================================================
    ' OBJETIVO:  Módulo de validação de CNPJ                 =
    ' DESENVOLVEDOR: Samuel Oliveira                         =
    ' ATUALIZAÇÃO: Conversão de letras A-Z para 10-35        =
    ' DATA: 2025-08-13                                       =
    '========================================================
    
    Dim CNPJ As String
    Dim validado As Boolean
    Dim posicao As Integer
    Dim ch As String
    Dim CalculoDigito As Long
    Dim strCNPJ As String
    Dim multiplicador As String
    Dim TamanhoCNPJ As Integer
    Dim etapa As Integer
    
    validado = False
    CNPJ = ""
    
    ' === Remover tudo que não for número ou letra ===
    For posicao = 1 To Len(numeroCNPJ)
        ch = Mid(numeroCNPJ, posicao, 1)
        If ch Like "[0-9A-Za-z]" Then
            ' Converter letras para números (A=10, B=11, ..., Z=35)
            If ch Like "[A-Za-z]" Then
                CNPJ = CNPJ & CStr(Asc(UCase(ch)) - 55)
            Else
                CNPJ = CNPJ & ch
            End If
        End If
    Next
    
    ' === CNPJ numérico final precisa ter 14 posições ===
    If Len(CNPJ) <> 14 Then Exit Function
    
    ' Lista de CNPJs inválidos
    Dim cnpjInvalidos As String
    cnpjInvalidos = "00000000000000,11111111111111,22222222222222,33333333333333,44444444444444," & _
                    "55555555555555,66666666666666,77777777777777,88888888888888,99999999999999"
    
    ' === Cálculo dos dígitos ===
    For etapa = 1 To 2
        multiplicador = IIf(etapa = 1, "543298765432", "6543298765432")
        TamanhoCNPJ = IIf(etapa = 1, 12, 13)
        
        For posicao = 1 To TamanhoCNPJ
            CalculoDigito = CalculoDigito + CLng(Mid(CNPJ, posicao, 1)) * CLng(Mid(multiplicador, posicao, 1))
        Next posicao
        
        CalculoDigito = IIf(CalculoDigito Mod 11 < 2, 0, 11 - (CalculoDigito Mod 11))
        strCNPJ = strCNPJ & CStr(CalculoDigito)
        CalculoDigito = 0
    Next etapa
    
    ' Comparar dígitos
    If CStr(Right(CNPJ, 2)) = strCNPJ Then validado = True
    
    ' Bloquear CNPJs inválidos conhecidos
    If InStr(1, cnpjInvalidos, CNPJ) > 0 Then validado = False
    
    validarCNPJ = validado
    Exit Function
    
TratarErro:
    MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbExclamation, "Erro " & Err.Number
End Function
