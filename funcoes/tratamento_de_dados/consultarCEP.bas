Function ConsultarCEP(CEP)
On Error GoTo TratarErro

'========================================================
' CONSULTA DE CEP                                       =
'========================================================
' OBJETIVO:         Módulo para consulta de CEPs        =
' DESENVOLVEDOR:    Samuel Oliveira                     =
' CONTATO:          samuel.santos@oprograma.com         =
' WEB SITE          www.oprograma.com                   =
' Copyright 2020 Todos os direitos reservados           =
'========================================================
' OBSERVAÇÕES                                           =
' É preciso adicionar a referência                      =
' "Microsoft XML, v3.0" ou superior                     =
'========================================================

'É preciso adicionar a referência "Microsoft XML, v3.0" ou superior

Dim oXmlDoc As DOMDocument
Dim oXmlNode As IXMLDOMNode
Dim oXmlNodes As IXMLDOMNodeList

Set DadosXML = New DOMDocument
DadosXML.async = False

DadosXML.Load ("https://viacep.com.br/ws/" + CEP + "/xml/")

NomeCampos = "cep;logradouro;complemento;bairro;localidade;uf"

ReDim ArrayCEP(1, UBound(Split(NomeCampos, ";")))

TemErro = UBound(Split(DadosXML.XML, "<erro>"))

For Etapa = 0 To 1
    For Linha = 0 To UBound(Split(NomeCampos, ";"))
    
        If Etapa = 0 Then ArrayCEP(Etapa, Linha) = Split(NomeCampos, ";")(Linha)
        If TemErro = 0 And Etapa = 1 Then ArrayCEP(Etapa, Linha) = Split(Split(DadosXML.XML, "</" & ArrayCEP(0, Linha) & ">")(0), "<" & ArrayCEP(0, Linha) & ">")(1)
           
    Next Linha
Next Etapa

ArrayCEP(1, 1) = Replace(ArrayCEP(1, 1), "Avenida", "Av.")

ConsultarCEP = ArrayCEP

SairFunction:
Exit Function

TratarErro:
MsgBox "Ocorreu um erro ao processar o comando:" & vbCrLf & Err.Description, vbCritical, " Erro " & Err.Number
Resume SairFunction
End Function
