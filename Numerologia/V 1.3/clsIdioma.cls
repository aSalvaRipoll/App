VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsIdioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' clsIdioma.cls
Option Compare Database
Option Explicit

Private mIDIdioma As Byte
Private mAbreviado As String
Private mNomIdioma As String
Private mNotas As String

' FUTURO: extensiones previstas
'Private mDialecto As String
'Private mRegion As String
'Private mReglasFoneticas As Object   ' Podrá ser una clase en el futuro
'Private mIconoAlternativo As String  ' Por si quieres banderas regionales


' -------------------------
' Inicialización
' -------------------------
Public Sub Init(ByVal ID As Byte, ByVal Abrev As String, _
                ByVal Nombre As String, Optional ByVal notas As String = "")

    mIDIdioma = ID
    mAbreviado = Abrev
    mNomIdioma = Nombre
    mNotas = notas

End Sub


' -------------------------
' Propiedades públicas
' -------------------------

Public Property Get IDIdioma() As Byte
    IDIdioma = mIDIdioma
End Property

Public Property Let IDIdioma(ByVal valor As Byte)
    mIDIdioma = valor
End Property


Public Property Get Abreviado() As String
    Abreviado = mAbreviado
End Property

Public Property Let Abreviado(ByVal valor As String)
    mAbreviado = valor
End Property


Public Property Get NomIdioma() As String
    NomIdioma = mNomIdioma
End Property

Public Property Let NomIdioma(ByVal valor As String)
    mNomIdioma = valor
End Property


Public Property Get notas() As String
    notas = mNotas
End Property

Public Property Let notas(ByVal valor As String)
    mNotas = valor
End Property


' -------------------------
' Propiedades futuras
' -------------------------

'Public Property Get Dialecto() As String
'    Dialecto = mDialecto
'End Property
'
'Public Property Let Dialecto(ByVal valor As String)
'    mDialecto = valor
'End Property
'
'
'Public Property Get Region() As String
'    Region = mRegion
'End Property
'
'Public Property Let Region(ByVal valor As String)
'    mRegion = valor
'End Property
'
'
'Public Property Get ReglasFoneticas() As Object
'    Set ReglasFoneticas = mReglasFoneticas
'End Property
'
'Public Property Set ReglasFoneticas(ByVal obj As Object)
'    Set mReglasFoneticas = obj
'End Property
'
'
'Public Property Get IconoAlternativo() As String
'    IconoAlternativo = mIconoAlternativo
'End Property
'
'Public Property Let IconoAlternativo(ByVal valor As String)
'    mIconoAlternativo = valor
'End Property


' -------------------------
' Métodos útiles
' -------------------------

Public Function RutaIcono() As String
    Dim ruta As String

'    If IconoAlternativo <> "" Then
'        'ruta = CurrentProject.Path & "\img\" & mIconoAlternativo & ".png"
'        ruta = CurrentProject.Path & "\img\" & IconoAlternativo & ".ico"
'    Else
        'ruta = CurrentProject.Path & "\img\" & mAbreviado & ".png"
        ruta = CurrentProject.Path & "\img\" & Abreviado & ".ico"
'    End If

    RutaIcono = ruta
End Function


