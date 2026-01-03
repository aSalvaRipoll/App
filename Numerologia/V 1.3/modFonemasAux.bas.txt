Attribute VB_Name = "modFonemasAux"

Option Compare Database
Option Explicit

Public Function EsVocal(ByVal c As String) As Boolean
    Select Case UCase$(c)
        Case "A", "Á", "À", _
             "E", "É", "È", _
             "I", "Í", "Ï", _
             "O", "Ó", "Ò", _
             "U", "Ú", "Ü"
            EsVocal = True
        Case Else
            EsVocal = False
    End Select
End Function

Public Function EsConsonante(ByVal c As String) As Boolean
    If c = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not EsVocal(c)
    End If
End Function

Public Function ProcesarY(ByVal ant As String, ByVal sig As String) As String
'Public Function ProcesarY(ByVal ant As String, ByVal c As String, ByVal sig As String) As String

    ' --- Tratamiento completo de la Y ---

    ' 1) Conjunción aislada: " y "
    If (ant = " ") And (sig = " ") Then
        ProcesarY = "I"
        Exit Function
    End If

    ' 2) Y inicial -> consonante si sig es vocal
    If ant = "" And EsVocal(sig) Then
        ProcesarY = "Y"
        Exit Function
    End If

    ' 3) Y inicial -> vocal si sig es consonante
    If ant = "" And EsConsonante(sig) Then
        ProcesarY = "I"
        Exit Function
    End If

    ' 4) Y final -> vocal (Rey ? Rei, Charly ? Charli)
    If sig = "" Then
        ProcesarY = "I"
        Exit Function
    End If

    ' 5) Y entre dos vocales -> consonante
    If EsVocal(ant) And EsVocal(sig) Then
        ProcesarY = "Y"
        Exit Function
    End If

    ' 6) Y entre dos consonantes -> vocal
    If EsConsonante(ant) And EsConsonante(sig) Then
        ProcesarY = "I"
        Exit Function
    End If

    ' 7) Y seguida de vocal -> consonante
    If EsVocal(sig) Then
        ProcesarY = "Y"
        Exit Function
    End If

    ' 8) Y seguida de consonante -> vocal
    If EsConsonante(sig) Then
        ProcesarY = "I"
        Exit Function
    End If

    ' Fallback seguro: tratar como consonante
    ProcesarY = "Y"

End Function

Public Function ProcesarW() As String
'Public Function ProcesarW(ByVal c As String) As String
    ' Opción A: la W siempre se pronuncia como GÜ
    ProcesarW = "GÜ"
End Function


' ============================================================================
' NORMALIZACIÓN OPTIMIZADA
'   - Mayúsculas
'   - Elimina tildes
'   - NO toca la Ü
' ============================================================================

'Private Function NormalizarTexto(ByVal Texto As String) As String
'    Dim i As Long
'    Dim c As String
'    Dim sb As String
'
'    Texto = UCase$(Texto)
'
'    For i = 1 To Len(Texto)
'        c = Mid$(Texto, i, 1)
'
'        Select Case c
'            Case "Á": sb = sb & "A"
'            Case "É": sb = sb & "E"
'            Case "Í": sb = sb & "I"
'            Case "Ó": sb = sb & "O"
'            Case "Ú": sb = sb & "U"
'            Case Else
'                sb = sb & c
'        End Select
'    Next i
'
'    NormalizarTexto = sb
'End Function



'Public Function ProcesarY( _
'    ByVal ant As String, _
'    ByVal c As String, _
'    ByVal sig As String _
'    ) As String
'
'    ' Devuelve: "Y", "I" o "" (no aplica)
'    ' Tu bloque tal cual, pero devolviendo el fonema o vacío si no se aplica
'     ' --- Tratamiento completo de la Y ---
'    If c = "Y" Then
'
'    ' 1) Conjunción aislada: " y "
'    If (ant = " ") And (sig = " ") Then
'        ProcesarY = "I"
'
'        Exit Function
'    End If
'
'    ' 2) Y inicial ? consonante si sig es vocal
'    If ant = "" And EsVocal(sig) Then
'        ProcesarY = "Y"
'
'        Exit Function
'    End If
'
'    ' 3) Y inicial ? vocal si sig NO es vocal
'    If ant = "" And Not EsVocal(sig) Then
'        ProcesarY = "I"
'
'        Exit Function
'    End If
'
'    ' 4) Y entre dos vocales ? consonante
'    If EsVocal(ant) And EsVocal(sig) Then
'        ProcesarY = "Y"
'
'        Exit Function
'    End If
'
'    ' 5) Y entre dos consonantes ? vocal
'    If Not EsVocal(ant) And Not EsVocal(sig) Then
'        ProcesarY = "I"
'
'        Exit Function
'    End If
'
'    ' 6) Y final tras vocal o consonante ? vocal (Rey ? Rei)
'    If sig = "" Then
'        ProcesarY = "I"
'
'        Exit Function
'    End If
'
'    ' 7) Y seguida de vocal ? consonante
'    If EsVocal(sig) Then
'        ProcesarY = "Y"
'
'        Exit Function
'    End If
'
'    ' 8) Y seguida de consonante ? vocal
'    If EsConsonante(sig) Then
'        ProcesarY = "I"
'
'        Exit Function
'    End If
'
'End Function
'

