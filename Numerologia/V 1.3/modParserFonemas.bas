Attribute VB_Name = "modParserFonemas"

Option Compare Database
Option Explicit

' ---------------------------------------------------------
'  Tipo de dato para un fonema alineado con la letra original
' ---------------------------------------------------------
Public Type tFonema
    LetraEscrita As String     ' Letra original del nombre
    FonemaASCII As String      ' Letra, espacio " " o fonema compuesto (CH, SH, NY, LY...)
    tipo As String             ' "V" = vocal, "C" = consonante, "" = ninguno (espacio)
    valor As Byte              ' Valor pitagorico (0 si no aplica)
End Type

' ---------------------------------------------------------
'  Array publico con el resultado del parser
' ---------------------------------------------------------
Public arrFonemas() As tFonema

Public Function ParseNombre(ByVal Nombre As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim i As Long
    Dim idx As Long
    Dim lenNombre As Long
    Dim ch As String
    Dim Siguiente As String
    
    Nombre = UCase$(Trim$(Nombre))
    lenNombre = Len(Nombre)
    
    If lenNombre = 0 Then
        ParseNombre = False
        Exit Function
    End If
    
    ReDim arrFonemas(1 To lenNombre)
    
    idx = 1
    i = 1
    
    Do While i <= lenNombre
        
        ch = Mid$(Nombre, i, 1)
        
        ' Inicializar
        arrFonemas(idx).LetraEscrita = ch
        arrFonemas(idx).FonemaASCII = ch
        arrFonemas(idx).tipo = TipoPorDefecto(ch)
        arrFonemas(idx).valor = ValorFonemaASCII(arrFonemas(idx).FonemaASCII)
        
        ' ---------------------------------------------------------
        '  DETECCION DE COMPUESTOS (orden correcto)
        ' ---------------------------------------------------------
        If i < lenNombre Then
            Siguiente = Mid$(Nombre, i + 1, 1)
            
            ' 1) TX --> CH
            If ch = "T" And Siguiente = "X" Then
                GoSub Procesar_TX
                GoTo SiguienteIteracion
            End If
            
            ' 2) CH --> CH
            If ch = "C" And Siguiente = "H" Then
                GoSub Procesar_CH
                GoTo SiguienteIteracion
            End If
            
            ' 3) SH --> SH
            If ch = "S" And Siguiente = "H" Then
                GoSub Procesar_SH
                GoTo SiguienteIteracion
            End If
            
            ' 4) NY --> NY
            If ch = "N" And Siguiente = "Y" Then
                GoSub Procesar_NY
                GoTo SiguienteIteracion
            End If
            
            ' 5) LY --> LY
            If ch = "L" And Siguiente = "Y" Then
                GoSub Procesar_LY
                GoTo SiguienteIteracion
            End If
        End If
        
        ' Avance normal
        i = i + 1
        idx = idx + 1
        
SiguienteIteracion:
    Loop
    
    ' Ajustar tamano real
    If idx > 1 Then
        ReDim Preserve arrFonemas(1 To idx - 1)
    End If
    
    ParseNombre = True
    Exit Function

' ---------------------------------------------------------
'  SUBRUTINAS DE COMPUESTOS
' ---------------------------------------------------------

Procesar_TX:
    arrFonemas(idx).FonemaASCII = " "
    arrFonemas(idx).tipo = ""
    arrFonemas(idx).valor = 0
    
    idx = idx + 1
    arrFonemas(idx).LetraEscrita = Siguiente
    arrFonemas(idx).FonemaASCII = "CH"
    arrFonemas(idx).tipo = "C"
    arrFonemas(idx).valor = ValorFonemaASCII("CH")
    
    i = i + 2
Return

Procesar_CH:
    arrFonemas(idx).FonemaASCII = " "
    arrFonemas(idx).tipo = ""
    arrFonemas(idx).valor = 0
    
    idx = idx + 1
    arrFonemas(idx).LetraEscrita = Siguiente
    arrFonemas(idx).FonemaASCII = "CH"
    arrFonemas(idx).tipo = "C"
    arrFonemas(idx).valor = ValorFonemaASCII("CH")
    
    i = i + 2
Return

Procesar_SH:
    arrFonemas(idx).FonemaASCII = " "
    arrFonemas(idx).tipo = ""
    arrFonemas(idx).valor = 0
    
    idx = idx + 1
    arrFonemas(idx).LetraEscrita = Siguiente
    arrFonemas(idx).FonemaASCII = "SH"
    arrFonemas(idx).tipo = "C"
    arrFonemas(idx).valor = ValorFonemaASCII("SH")
    
    i = i + 2
Return

Procesar_NY:
    arrFonemas(idx).FonemaASCII = " "
    arrFonemas(idx).tipo = ""
    arrFonemas(idx).valor = 0
    
    idx = idx + 1
    arrFonemas(idx).LetraEscrita = Siguiente
    arrFonemas(idx).FonemaASCII = "NY"
    arrFonemas(idx).tipo = "C"
    arrFonemas(idx).valor = ValorFonemaASCII("NY")
    
    i = i + 2
Return

Procesar_LY:
    arrFonemas(idx).FonemaASCII = " "
    arrFonemas(idx).tipo = ""
    arrFonemas(idx).valor = 0
    
    idx = idx + 1
    arrFonemas(idx).LetraEscrita = Siguiente
    arrFonemas(idx).FonemaASCII = "LY"
    arrFonemas(idx).tipo = "C"
    arrFonemas(idx).valor = ValorFonemaASCII("LY")
    
    i = i + 2
Return

ErrHandler:
    ParseNombre = False
End Function

Private Function ValorFonemaASCII(ByVal fonema As String) As Byte
    Select Case fonema
    
        Case "CH": ValorFonemaASCII = 3
        Case "SH": ValorFonemaASCII = 1
        Case "NY": ValorFonemaASCII = 5
        Case "LY": ValorFonemaASCII = 3
    
        Case "A": ValorFonemaASCII = 1
        Case "E": ValorFonemaASCII = 5
        Case "I": ValorFonemaASCII = 9
        Case "O": ValorFonemaASCII = 6
        Case "U": ValorFonemaASCII = 3
        
        Case "B": ValorFonemaASCII = 2
        Case "C": ValorFonemaASCII = 3
        Case "D": ValorFonemaASCII = 4
        Case "F": ValorFonemaASCII = 6
        Case "G": ValorFonemaASCII = 7
        Case "L": ValorFonemaASCII = 3
        Case "M": ValorFonemaASCII = 4
        Case "N": ValorFonemaASCII = 5
        Case "P": ValorFonemaASCII = 7
        Case "R": ValorFonemaASCII = 9
        Case "S": ValorFonemaASCII = 1
        Case "T": ValorFonemaASCII = 2
        Case "Z": ValorFonemaASCII = 8
        
        Case " "
            ValorFonemaASCII = 0
        Case Else
            ValorFonemaASCII = 0
    End Select
End Function

