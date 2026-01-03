VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsConversorFonetico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

' ============================================================================
' Clase: clsConversorFonetico
' Descripción: Convierte un nombre a su representación fonética según idioma
' ============================================================================

' Propiedades privadas
Private pNombreOriginal As String
Private pNombreFonetico As String
Private pFonemas As Collection
Private pIdioma As idioma
Private pUsarHmuda As Boolean
Private pUsarUmuda As Boolean

' ============================================================================
' PROPIEDADES PÚBLICAS
' ============================================================================

Public Property Let NombreOriginal(ByVal valor As String)
    pNombreOriginal = Trim$(valor)
End Property

Public Property Get NombreOriginal() As String
    NombreOriginal = pNombreOriginal
End Property


Public Property Get NombreFonetico() As String
    NombreFonetico = pNombreFonetico
End Property


Public Property Get Fonemas() As Collection
    Set Fonemas = pFonemas
End Property


Public Property Let idioma(ByVal valor As idioma)
    pIdioma = valor
End Property

Public Property Get idioma() As idioma
    idioma = pIdioma
End Property


Public Property Let UsarHmuda(ByVal valor As Boolean)
    pUsarHmuda = valor
End Property

Public Property Let UsarUmuda(ByVal valor As Boolean)
    pUsarUmuda = valor
End Property

' ============================================================================
' MÉTODO PRINCIPAL
' ============================================================================

Public Function Convertir() As String
    Dim col As Collection
    Dim f As Variant
    Dim sb As String
    
    If Len(pNombreOriginal) = 0 Then
        Err.Raise vbObjectError + 3001, "clsConversorFonetico", "El nombre está vacío"
    End If
    
    ' Selección del tokenizador según idioma
    Select Case pIdioma
    
        Case iEspañol
            Set col = ObtenerFonemasCastellano(pNombreOriginal, pUsarHmuda, pUsarUmuda)
        
        Case iCatala
            Set col = ObtenerFonemasCatalan(pNombreOriginal, pUsarHmuda, pUsarUmuda)
        
        Case iEuskera
            Set col = ObtenerFonemasEuskera(pNombreOriginal)
        
        Case iGalego
            Set col = ObtenerFonemasGalego(pNombreOriginal, pUsarHmuda, pUsarUmuda)
        
        Case Else
            Err.Raise vbObjectError + 3005, "clsConversorFonetico", "Idioma no válido"
    End Select
    
    ' Guardar fonemas
    Set pFonemas = col
    
    ' Construir nombre fonético concatenado
    sb = ""
    For Each f In col
        sb = sb & f
    Next f
    
    pNombreFonetico = sb
    Convertir = sb
End Function

'' ============================================================================
'' MÉTODO DE DEPURACIÓN
'' ============================================================================
'
'Public Function ObtenerInformacion() As String
'    Dim info As String
'    Dim f As Variant
'

