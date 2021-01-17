Function LerPropriedadesArquivo(NomeArquivo)
On Error GoTo TratarErro

If IsNull(IIf(NomeArquivo = "", Null, NomeArquivo)) Then Exit Function

If Not IsNull(IIf(NomeArquivo = "", Null, NomeArquivo)) Then

    ReDim MeuArray(1, 6)

    Set Objeto = CreateObject("Scripting.FileSystemObject")
    Set Arquivo = Objeto.GetFile(NomeArquivo)
    
    TipoArquivo = Arquivo.Type
    DataCriado = Arquivo.DateCreated
    DataUltimoAcesso = Arquivo.DateLastAccessed
    DataUltimaModificacao = Arquivo.DateLastModified
    Tamanho = Round((Arquivo.Size / 1024), 2) & " KB"
    
    Set Arquivo = LoadPicture(NomeArquivo)
    
    TamanhoWidth = Round(Arquivo.Width / 26.4583)
    TamanhoHeight = Round(Arquivo.Height / 26.4583)
    
    MeuArray(0, 0) = "Tipo de Arquivo"
    MeuArray(0, 1) = "Data Criação"
    MeuArray(0, 2) = "Último Acesso"
    MeuArray(0, 3) = "Última modificação"
    MeuArray(0, 4) = "Tamanho KB"
    MeuArray(0, 5) = "Width"
    MeuArray(0, 6) = "Height"
    
    MeuArray(1, 0) = TipoArquivo
    MeuArray(1, 1) = DataCriado
    MeuArray(1, 2) = DataUltimoAcesso
    MeuArray(1, 3) = DataUltimaModificacao
    MeuArray(1, 4) = Tamanho
    MeuArray(1, 5) = TamanhoWidth
    MeuArray(1, 6) = TamanhoHeight
    
    Set Objeto = Nothing

SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbCritical, " Erro " & Err.Number
Resume SairFunction

End If
