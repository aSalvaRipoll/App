VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCatalogoInspector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'=====================================================
' Clase: clsCatalogoInspector
' Catálogo completo del proyecto VBA
'=====================================================

Option Compare Database
Option Explicit

Public Modulos As Collection
Public Clases As Collection
Public UserForms As Collection
Public Formularios As Collection
Public Informes As Collection
Public Otros As Collection

Private Sub Class_Initialize()
    Set Modulos = New Collection
    Set Clases = New Collection
    Set UserForms = New Collection
    Set Formularios = New Collection
    Set Informes = New Collection
    Set Otros = New Collection
End Sub

