VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConversorFonético"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public Function NombreFoneticoDesdeTexto(ByVal nombre As String, _
                                         ByVal Idioma As Idioma, _
                                         Optional ByVal UsarHmuda As Boolean = True, _
                                         Optional ByVal UsarUmuda As Boolean = True) As String
    Dim colFonemas As Collection
    Dim sb As String
    Dim f As Variant
    
    Select Case Idioma
        Case iEspañol
            Set colFonemas = ObtenerFonemasCastellano(nombre, UsarHmuda, UsarUmuda)
        ' Más adelante: iCatala, iEuskera, iGalego...
        Case Else
            Err.Raise vbObjectError + 2000, "clsConversorFonético", "Idioma no soportado todavía"
    End Select
    
    For Each f In colFonemas
        ' Separador opcional, por ejemplo un espacio o punto
        sb = sb & f
    Next f
    
    NombreFoneticoDesdeTexto = sb
End Function


