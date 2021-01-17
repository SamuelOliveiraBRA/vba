Function apenasNumero(texto)

For posicao = 1 To Len(texto)
    'If InStr(1, numeros, Mid(texto, posicao, 1)) > 0 Then
    If Mid(texto, posicao, 1) Like "[0-9]" Then
        novoTexto = novoTexto & Mid(texto, posicao, 1)
    End If
Next posicao

apenasNumero = novoTexto

End Function
