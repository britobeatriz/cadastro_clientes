VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormPrincipal 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   10692
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17724
   OleObjectBlob   =   "FormPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BTBuscar_Click()
    FormBuscar.Show
End Sub

Private Sub BTEditar_Click()
    FormEditar.Show
End Sub

Private Sub BTExcluir_Click()
    FormExcluir.Show
End Sub

Private Sub BTIncluir_Click()
    FormIncluir.Show
End Sub


