Attribute VB_Name = "webscrap"
Option Explicit
Sub scrapApi()
    'Pesquisar o cep
    
    Dim api As New MSXML2.ServerXMLHTTP60
    Dim html As New MSHTML.HTMLDocument
    Dim url As String
    Dim cep As String
    
    cep = FormIncluir.TextCep.Value
    url = "https://viacep.com.br/ws/" & cep & "/xml/"
    
    api.Open "GET", url
    api.send
    
    On Error GoTo invalido
    
    html.body.innerHTML = api.responseText
   
    FormIncluir.TextRua.Value = html.getElementsByTagName("logradouro")(0).innerText
    FormIncluir.TextBairro.Value = html.getElementsByTagName("bairro")(0).innerText
    FormIncluir.TextEstado.Value = html.getElementsByTagName("uf")(0).innerText
    FormIncluir.TextCidade.Value = html.getElementsByTagName("localidade")(0).innerText
    Exit Sub
    
invalido:
    FormIncluir.TextRua.Value = ""
    FormIncluir.TextBairro.Value = ""
    FormIncluir.TextEstado.Value = ""
    FormIncluir.TextCidade.Value = ""
    MsgBox "CEP inválido"
    
End Sub
