Attribute VB_Name = "modMotor_Idioma_ES"

Option Compare Database
Option Explicit

'=================
'==  Castellano ==
'=================



'=======================================================================================


' ============================================================
'   Silabear_ES — Silabeador para nombres y apellidos en español
'   - Respeta espacios entre palabras
'   - No mezcla sílabas entre palabras
'   - Aplica reglas fonéticas del español
'   - Detecta dígrafos (CH, LL, RR)
'   - Detecta grupos consonánticos inseparables (BR, CR, TR…)
'   - Trata la H como consonante muda (no rompe sílabas)
'   - Elimina H final aislada
'   - Devuelve posiciones absolutas (ini, fin)
' ============================================================

' ============================================================
'   REVISIÓN MANUAL MEDIANTE INPUTBOX
' ============================================================

Public Function Silabear_ES_ConRevision(ByVal Texto As String) As Collection

    Dim col As Collection
    Dim resultado As New Collection
    Dim item As Variant
    Dim s As String
    Dim partes As Variant
    Dim p As Variant
    'Dim inicio As Long, fin As Long
    Dim valido As Boolean
    Dim msg As String
    Dim i As Long
    
    Dim ini As Long, fin As Long

    Dim palabras() As String
    Dim sils() As String
    Dim s2 As Variant
    Dim pos As Long

    ' 1. Silabear automáticamente
    Set col = Silabear_ES(Texto)

    ' 2. Convertir a string con separador visual "-"
    '    (las sílabas no incluyen espacios; los espacios están en Texto)
    s = ""
    
    For i = 1 To col.Count
    
        ' Añadir la sílaba actual
        s = s & Mid$(Texto, col(i)(0), col(i)(1) - col(i)(0) + 1)
    
        ' Si no es la última sílaba
        If i < col.Count Then
    
            ' Si hay un espacio entre esta sílaba y la siguiente
            If col(i + 1)(0) > col(i)(1) + 1 Then
                s = s & " "   ' espacio real
            Else
                s = s & "-"   ' separador silábico
            End If
    
        End If
    
    Next i

'    ' 2. Convertir a string con separador visual "-"
'    '    (las sílabas no incluyen espacios; los espacios están en Texto)
'    For Each item In col
'        s = s & Mid$(texto, item(0), item(1) - item(0) + 1) & "-"
'    Next item
'    If Len(s) > 0 Then s = Left$(s, Len(s) - 1)

    ' ============================================================
    ' 3. Bucle de validación
    ' ============================================================
    Do
        valido = True
        msg = ""

        's = InputBox("Revisa o corrige las sílabas:" & vbCrLf & _
                     "(usa '-' como separador entre sílabas)", _
                     "Revisión de sílabas", s)


        s = RevisarSilabas_EnFormulario(Texto, s)


        ' Si el usuario cancela --> devolver silabeo automático
        If s = "" Then
            Set Silabear_ES_ConRevision = col
            Exit Function
        End If

        ' No recortamos espacios: pueden ser parte del nombre compuesto
        ' pero sí limpiamos dobles guiones accidentales

        ' Validación 1: no puede empezar ni acabar con "-"
        If Left$(s, 1) = "-" Or Right$(s, 1) = "-" Then
            valido = False
            msg = "No puede empezar ni terminar con '-'."
        End If

        ' Validación 2: no puede contener "--" (sílabas vacías)
        If InStr(s, "--") > 0 Then
            valido = False
            msg = "No puede haber sílabas vacías ('--')."
        End If

        ' Validación 3: comprobar que las sílabas reconstruyen el texto original
        ' Ignorando espacios y separadores
        Dim reconstruido As String
        Dim textoSinEspacios As String

        ' Quitar separadores "-" y espacios del string de sílabas
        reconstruido = Replace(s, "-", "")
        reconstruido = Replace(reconstruido, " ", "")

        ' Quitar espacios del texto original
        textoSinEspacios = Replace(Texto, " ", "")

        If UCase$(reconstruido) <> UCase$(textoSinEspacios) Then
            valido = False
            msg = "Las sílabas no coinciden con el texto original (ignorando espacios)."
        End If

        ' Si no es válido --> mostrar mensaje y repetir
        If Not valido Then
            MsgBox msg, vbExclamation, "Error en las sílabas"
        End If

    Loop Until valido

'    ' ============================================================
'    ' 4. Reconstruir colección válida
'    ' ============================================================
'    partes = Split(s, "-")
'    inicio = 1
'
'    For Each p In partes
'        fin = inicio + Len(p) - 1
'        resultado.Add Array(inicio, fin)
'        inicio = fin + 1
'    Next p

' ============================================================
' 4. Reconstruir colección válida (respetando espacios)
' ============================================================

Set resultado = New Collection

' Posición absoluta dentro del texto original
pos = 1

' Dividir por palabras (separadas por espacios)
palabras = Split(s, " ")

For Each p In palabras

    ' Dividir cada palabra en sílabas
    sils = Split(p, "-")

    For Each s2 In sils

        ' Saltar sílabas vacías (por seguridad)
        If Trim$(s2) <> "" Then

            ini = InStr(pos, Texto, s2)

            If ini = 0 Then
                ' No encontrado: error de alineación
                MsgBox "Error: la sílaba '" & s2 & "' no se encuentra en el texto original.", vbCritical
                Exit Function
            End If

            fin = ini + Len(s2) - 1

            resultado.Add Array(ini, fin)

            ' Avanzar la posición de búsqueda
            pos = fin + 1

        End If

    Next s2

    ' Después de cada palabra, avanzar un espacio si existe
    If pos <= Len(Texto) Then
        If Mid$(Texto, pos, 1) = " " Then
            pos = pos + 1
        End If
    End If

Next p

    Set Silabear_ES_ConRevision = resultado

End Function








'================================================================================


'=========================================================================================

'============================================================================================


