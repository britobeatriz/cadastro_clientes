Attribute VB_Name = "webscrap"
Option Explicit
Sub scrapApi()
    'Pesquisar o cep
    
    Dim req As New XMLHTTP60
    Dim cep As String
    
    cep = FormIncluir.TextCep
    
     
        
With req
    .Open "GET", "https://viacep.com.br/ws/" & cep & "/xml/"
    .send
End With

On Error GoTo invalido
    
With Application.WorksheetFunction
    FormIncluir.TextRua.Value = .FilterXML(req.responseText, "//xmlcep/logradouro")
    FormIncluir.TextBairro.Value = .FilterXML(req.responseText, "//xmlcep/bairro")
    FormIncluir.TextEstado.Value = .FilterXML(req.responseText, "//xmlcep/uf")
    FormIncluir.TextCidade.Value = .FilterXML(req.responseText, "//xmlcep/localidade")
End With

      
   Exit Sub
    
invalido:
    FormIncluir.TextRua.Value = ""
    FormIncluir.TextBairro.Value = ""
    FormIncluir.TextEstado.Value = ""
    FormIncluir.TextCidade.Value = ""
    MsgBox "CEP inválido"
    
End Sub
