Attribute VB_Name = "modMotorFonetico_V2_1_Aux"

Option Compare Database
Option Explicit

Public Function RevisarSilabas_EnFormulario( _
        ByVal texto As String, _
        ByVal silabas As String _
    ) As String

    DoCmd.OpenForm "frmRevisionSilabas", , , , , acHidden ', WindowMode:=acDialog

    With Forms!frmRevisionSilabas
        .TextoOriginal = texto
        .SilabasOriginal = silabas
        .lblOriginal.Caption = texto
        '.txtSilabasOriginal.Value = .ResaltarTonicaHTML(Silabas) '.InsertarEspaciosEnSilabas(texto, Silabas)
        .txtSilabasOriginal.Value = .NormalizarYResaltarTonica(silabas)
        .txtSilabas = .txtSilabasOriginal
        
        .Visible = True
    End With

    ' Espera a que el formulario se cierre
    'Do While CurrentProject.AllForms("frmRevisionSilabas").IsLoaded
    Do While Forms!frmRevisionSilabas.Visible
        DoEvents
    Loop

    If Forms!frmRevisionSilabas.Cancelado Then
        RevisarSilabas_EnFormulario = silabas
    Else
        RevisarSilabas_EnFormulario = Forms!frmRevisionSilabas.SilabasFinal
    End If
    
    ' Cerrar el formulario aquí, cuando ya hemos leído los datos
    DoCmd.Close acForm, "frmRevisionSilabas"
    
End Function


Public Function MF_SilabearUniversalBase(ByVal texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long
    
    Dim vocales As String
    Dim vocalesDebiles As String
    Dim vocalesFuertes As String
    
    Dim esV1 As Boolean, esV2 As Boolean
    
    Dim c1 As String, c2 As String
    Dim diptongo As Boolean
    
    ' Vocales universales (romance + francés + portugués + catalán + euskara + robustez)
    
'    vocales = "AEIOU" & _
              "ÁÉÍÓÚ" & _
              "ÀÈÌÒÙ" & _
              "ÂÊÔ" & _
              "ÃÕ" & _
              "Ü" & _
              "Ÿ"

    vocales = "AEIOU" & _
        ChrW(&HC1) & ChrW(&HC9) & ChrW(&HCD) & ChrW(&HD3) & ChrW(&HDA) & _
        ChrW(&HC0) & ChrW(&HC8) & ChrW(&HCC) & ChrW(&HD2) & ChrW(&HD9) & _
        ChrW(&HC2) & ChrW(&HCA) & ChrW(&HD4) & _
        ChrW(&HC3) & ChrW(&HD5) & _
        ChrW(&HDC) & _
        ChrW(&H178)  ' ÁÉÍÓÚ ' ÀÈÌÒÙ ' ÂÊÔ ' ÃÕ ' Ü

    
    ' Vocales débiles universales
    
'    vocalesDebiles = "IÍÌUÚÙÜŸ"

'    vocalesDebiles = "IU" & _
        ChrW(&HCD) & ChrW(&HCC) & _
        ChrW(&HDA) & ChrW(&HD9) & _
        ChrW(&HDC) & _
        ChrW(&H178) ' Í Ì ' Ú Ù ' Ü ' Ÿ

    vocalesDebiles = "IU" & ChrW(&HDC)   ' I, U, Ü

    ' Vocales fuertes universales
    
'    vocalesFuertes = "AÁÀÂÃEÉÈÊOÓÒÔÕ"

'    vocalesFuertes = "AEO" & _
        ChrW(&HC1) & ChrW(&HC0) & ChrW(&HC2) & ChrW(&HC3) & _
        ChrW(&HC9) & ChrW(&HC8) & ChrW(&HCA) & _
        ChrW(&HD3) & ChrW(&HD2) & ChrW(&HD4) & ChrW(&HD5)    ' Á À Â Ã ' É È Ê ' Ó Ò Ô Õ

    vocalesFuertes = "AEO" & _
        ChrW(&HC1) & ChrW(&HC0) & ChrW(&HC2) & ChrW(&HC3) & _
        ChrW(&HC9) & ChrW(&HC8) & ChrW(&HCA) & _
        ChrW(&HD3) & ChrW(&HD2) & ChrW(&HD4) & ChrW(&HD5) & _
        ChrW(&HCD) & ChrW(&HCC) & ChrW(&HDA) & ChrW(&HD9) & _
        ChrW(&H178)

    ini = 1

    For i = 2 To Len(texto)
        
        c1 = Mid$(texto, i - 1, 1)
        c2 = Mid$(texto, i, 1)
        
        esV1 = (InStr(vocales, c1) > 0)
        esV2 = (InStr(vocales, c2) > 0)

        ' --------------------------------------------------------
        ' Caso 1: Vocal + Vocal --> decidir diptongo o hiato
        ' --------------------------------------------------------
        If esV1 And esV2 Then
            
            diptongo = False

            ' Diptongo universal:
            ' fuerte+débil, débil+fuerte, débil+débil
            If (InStr(vocalesFuertes, c1) > 0 And InStr(vocalesDebiles, c2) > 0) _
                Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesFuertes, c2) > 0) _
                Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesDebiles, c2) > 0) Then
                diptongo = True
            End If

            ' Si NO hay diptongo --> cerrar sílaba
            If Not diptongo Then
                col.Add Array(ini, i - 1)
                ini = i
            End If

        End If
    
    Debug.Print
    Debug.Print "vocales: "; vocales
    Debug.Print "vocalesDebiles: "; vocalesDebiles
    Debug.Print "vocalesFuertes: "; vocalesFuertes
    Debug.Print "c1: "; c1; " , c2: "; c2
    Debug.Print "esV1: "; esV1; " , esV2: "; esV2
    Debug.Print "diptongo: "; diptongo; " ,  i: "; i
    
    Next i

'Debug.Print "=== Silabas iniciales ==="
'Dim k As Integer
'For k = 1 To silabas.Count
'    Debug.Print "Silaba"; k; _
'                "Tipo=" & TypeName(silabas(k)); _
'                "Len=" & IIf(IsArray(silabas(k)), UBound(silabas(k)), "NO ARRAY")
'Next k

    ' Última sílaba
    col.Add Array(ini, Len(texto))

    Set MF_SilabearUniversalBase = col

End Function

'-----------------------------------------------------------------------------------
Public Function esVocal(ByVal c As String) As Boolean
    esVocal = (InStr("AEIOUÁÉÍÓÚÀÈÌÒÙÂÊÔÃÕÜŸ", c) > 0)
End Function


Public Function EsConsonante(ByVal c As String) As Boolean
    If c = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not esVocal(c)
    End If
End Function

' ------------------------------------------------------------
' 1. EsMiembro: comprueba si una cadena está en un array
' ------------------------------------------------------------
Public Function EsMiembro(ByVal s As String, ByVal arr As Variant) As Boolean
    Dim x As Variant
    For Each x In arr
        If s = x Then
            EsMiembro = True
            Exit Function
        End If
    Next x
    EsMiembro = False
End Function

' ------------------------------------------------------------
' 3. MF_UnirConsonantesEnAtaque
'    Une dos sílabas cuando el universal las separó
' ------------------------------------------------------------
Public Sub MF_UnirConsonantesEnAtaque( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        ' Si la división está justo antes de pos --> unir
        If fin = pos - 1 Then
            Dim ini2 As Long, fin2 As Long
            ini2 = silabas(i + 1)(0)
            fin2 = silabas(i + 1)(1)

            silabas.Remove i + 1
            silabas.Remove i
            silabas.Add Array(ini, fin2), , i
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 4. MF_UnirVocalesEnDiptongo
'    Une dos sílabas cuando deben formar diptongo
' ------------------------------------------------------------
Public Sub MF_UnirVocalesEnDiptongo( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        ' Si la sílaba empieza en pos --> unir con la anterior
        If ini = pos Then
            Dim iniPrev As Long, finPrev As Long
            iniPrev = silabas(i - 1)(0)
            finPrev = silabas(i - 1)(1)

            silabas.Remove i
            silabas.Remove i - 1
            silabas.Add Array(iniPrev, fin), , i - 1
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 5. MF_ForzarDivisionSilabica
'    Divide una sílaba en dos en un punto concreto
' ------------------------------------------------------------
Public Sub MF_ForzarDivisionSilabica( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        If pos > ini And pos <= fin Then
            ' Dividir esta sílaba en dos
            silabas.Remove i
            silabas.Add Array(ini, pos - 1), , i
            silabas.Add Array(pos, fin), , i + 1
            Exit Sub
        End If
    Next i

End Sub

' ------------------------------------------------------------
' 6. MF_EliminarVocalFinal
'    (para francés: elimina la E muda final)
' ------------------------------------------------------------
Public Sub MF_EliminarVocalFinal( _
        ByRef silabas As Collection, _
        ByVal pos As Long _
    )

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        If fin = pos Then
            silabas.Remove i
            silabas.Add Array(ini, pos - 1), , i
            Exit Sub
        End If
    Next i

End Sub

Public Sub MF_MarcarTonicaPenultima( _
        ByVal texto As String, _
        ByRef esTonica() As Boolean)

    Dim i As Long
    Dim ultimaVocal As Long
    Dim penultimaVocal As Long

    ' Buscar vocales
    For i = Len(texto) To 1 Step -1
        If esVocal(Mid$(texto, i, 1)) Then
            If ultimaVocal = 0 Then
                ultimaVocal = i
            Else
                penultimaVocal = i
                Exit For
            End If
        End If

    Next i

    ' Si hay penúltima vocal --> tónica
    If penultimaVocal > 0 Then
        esTonica(penultimaVocal) = True
    ElseIf ultimaVocal > 0 Then
        esTonica(ultimaVocal) = True
    End If

End Sub

Public Function MF_NormalizarVocales_General(ByVal texto As String) As String
    
    Dim i As Integer
    Dim c As String
    Dim res As String

    c = ""
    res = ""
    
    For i = 1 To Len(texto)
        c = Mid(texto, i, 1)
        Select Case c
            
            Case "Á", "À", "Ä"
                c = "A"
    
            Case "É", "È", "Ë"
                c = "E"
    
            Case "Í", "Ï"
                c = "I"
            
            Case "Ó", "Ò", "Ö"
                c = "O"
    
            Case "Ú", "Ü"
                c = "U"
        End Select
        
        res = res & c
    Next i
    
    MF_NormalizarVocales_General = res

End Function


Public Function ProcesarY(ByVal ant As String, ByVal sig As String) As String
    If (ant = " ") And (sig = " ") Then
        ProcesarY = "I": Exit Function
    End If

    If ant = "" And esVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If ant = "" And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If sig = "" Then
        ProcesarY = "I": Exit Function
    End If

    If esVocal(ant) And esVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If EsConsonante(ant) And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If esVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    ProcesarY = "Y"
End Function

Public Function ProcesarW() As String
    ProcesarW = "GÜ"
End Function

Public Sub CargarFonemas()
    Dim rs As DAO.Recordset
    Dim f As clsFonema

    If Not colFonemas Is Nothing Then Exit Sub   ' Ya cargada

    Set colFonemas = New Collection
    'Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbmFonemas ORDER BY idFonema")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbmFoneticaCompleta ORDER BY idFonema")

    Do While Not rs.EOF
        Set f = New clsFonema
        f.idFonema = rs!idFonema
        f.GrafemaOri = rs!fonema
        f.esVocal = rs!esVocal
        'f.Valor = rs!Valor
        ' ... cualquier otra propiedad

        colFonemas.Add f, CStr(f.idFonema)
        rs.MoveNext
    Loop

    rs.Close
End Sub

'-----------------------------------------------------------------------------------

'Public Function EsVocal(ByVal c As String) As Boolean
'    EsVocal = InStr("AEIOUÁÉÍÓÚÀÈÌÒÙÂÊÔÃÕÜ", c) > 0
'End Function

'Public Function MF_SilabearUniversalBase(ByVal Texto As String) As Collection
'
'    Dim col As New Collection
'    Dim i As Long, ini As Long
'
'    ' Vocales universales (romance + francés + portugués + catalán + euskara)
'    Dim vocales As String
'    vocales = "AEIOU" & _
'              "ÁÉÍÓÚ" & _
'              "ÀÈÌÒÙ" & _
'              "ÂÊÔ" & _
'              "ÃÕ" & _
'              "Ü"
'
'    ' Vocales débiles universales
'    Dim vocalesDebiles As String
'    vocalesDebiles = "IÍÌUÚÙÜ"
'
'    ' Vocales fuertes universales
'    Dim vocalesFuertes As String
'    vocalesFuertes = "AÁÀÂÃEÉÈÊOÓÒÔÕ"
'
'    ini = 1
'
'    For i = 2 To Len(Texto)
'
'        Dim c1 As String, c2 As String
'        c1 = Mid$(Texto, i - 1, 1)
'        c2 = Mid$(Texto, i, 1)
'
'        Dim esV1 As Boolean, esV2 As Boolean
'        esV1 = (InStr(vocales, c1) > 0)
'        esV2 = (InStr(vocales, c2) > 0)
'
'        ' --------------------------------------------------------
'        ' Caso 1: Vocal + Vocal --> decidir diptongo o hiato
'        ' --------------------------------------------------------
'        If esV1 And esV2 Then
'
'            Dim diptongo As Boolean
'            diptongo = False
'
'            ' Diptongo universal:
'            ' fuerte+débil, débil+fuerte, débil+débil
'            If (InStr(vocalesFuertes, c1) > 0 And InStr(vocalesDebiles, c2) > 0) _
'            Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesFuertes, c2) > 0) _
'            Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesDebiles, c2) > 0) Then
'                diptongo = True
'            End If
'
'            ' Si NO hay diptongo --> cerrar sílaba
'            If Not diptongo Then
'                col.Add Array(ini, i - 1)
'                ini = i
'            End If
'
'        End If
'
'        ' --------------------------------------------------------
'        ' Caso 2: Consonante --> Vocal
'        ' (posible inicio de sílaba, pero no cerramos aún)
'        ' --------------------------------------------------------
'        ' No hacemos nada aquí
'
'        ' --------------------------------------------------------
'        ' Caso 3: Vocal --> Consonante
'        ' (posible cierre, pero esperamos a la siguiente vocal)
'        ' --------------------------------------------------------
'        ' No hacemos nada aquí
'
'    Next i
'
'    ' Última sílaba
'    col.Add Array(ini, Len(Texto))
'
'    Set MF_SilabearUniversalBase = col
'
'End Function



'Private Sub MF_EliminarVocalFinal(ByRef silabas As Collection, ByVal pos As Long)
'
'    Dim i As Long
'
'    For i = 1 To silabas.Count
'        Dim ini As Long, fin As Long
'        ini = silabas(i)(0)
'        fin = silabas(i)(1)
'
'        If fin = pos Then
'            ' Reducir la sílaba eliminando la E final
'            silabas.Remove i
'            silabas.Add Array(ini, pos - 1), , i
'            Exit Sub
'        End If
'    Next i
'
'End Sub

'Private Sub MF_ForzarDivisionSilabica(ByRef silabas As Collection, ByVal pos As Long)
'
'    Dim i As Long
'
'    For i = 1 To silabas.Count
'        Dim ini As Long, fin As Long
'        ini = silabas(i)(0)
'        fin = silabas(i)(1)
'
'        If pos > ini And pos <= fin Then
'            ' Dividir esta sílaba en dos
'            silabas.Remove i
'            silabas.Add Array(ini, pos - 1), , i
'            silabas.Add Array(pos, fin), , i + 1
'            Exit Sub
'        End If
'    Next i
'
'End Sub

'Private Sub MF_UnirVocalesEnDiptongo(ByRef silabas As Collection, ByVal pos As Long)
'
'    Dim i As Long
'
'    For i = 1 To silabas.Count
'        Dim ini As Long, fin As Long
'        ini = silabas(i)(0)
'        fin = silabas(i)(1)
'
'        If ini = pos Then
'            Dim iniPrev As Long, finPrev As Long
'            iniPrev = silabas(i - 1)(0)
'            finPrev = silabas(i - 1)(1)
'
'            silabas.Remove i
'            silabas.Remove i - 1
'            silabas.Add Array(iniPrev, fin), , i - 1
'            Exit Sub
'        End If
'    Next i
'
'End Sub
'
'Private Function EsMiembroCA(ByVal s As String, ByVal arr As Variant) As Boolean
'    Dim x As Variant
'    For Each x In arr
'        If s = x Then
'            EsMiembro = True
'            Exit Function
'        End If
'    Next x
'    EsMiembro = False
'End Function


'Private Sub MF_UnirConsonantesEnAtaque(ByRef silabas As Collection, ByVal pos As Long)
'
'    Dim i As Long
'
'    For i = 1 To silabas.Count
'        Dim ini As Long, fin As Long
'        ini = silabas(i)(0)
'        fin = silabas(i)(1)
'
'        ' Si la división está justo antes de pos --> unir
'        If fin = pos - 1 Then
'            ' Unir esta sílaba con la siguiente
'            Dim ini2 As Long, fin2 As Long
'            ini2 = silabas(i + 1)(0)
'            fin2 = silabas(i + 1)(1)
'
'            silabas.Remove i + 1
'            silabas.Remove i
'            silabas.Add Array(ini, fin2), , i
'            Exit Sub
'        End If
'    Next i
'
'End Sub
'
'Private Sub MF_UnirConsonantesEnAtaqueCA(ByRef silabas As Collection, ByVal pos As Long)
'
'    Dim i As Long
'
'    For i = 1 To silabas.Count
'        Dim ini As Long, fin As Long
'        ini = silabas(i)(0)
'        fin = silabas(i)(1)
'
'        If fin = pos - 1 Then
'            Dim ini2 As Long, fin2 As Long
'            ini2 = silabas(i + 1)(0)
'            fin2 = silabas(i + 1)(1)
'
'            silabas.Remove i + 1
'            silabas.Remove i
'            silabas.Add Array(ini, fin2), , i
'            Exit Sub
'        End If
'    Next i
'
'End Sub

'Private Function EsMiembro(ByVal s As String, ByVal arr As Variant) As Boolean
'    Dim x As Variant
'    For Each x In arr
'        If s = x Then
'            EsMiembro = True
'            Exit Function
'        End If
'    Next x
'    EsMiembro = False
'End Function

'Public Function Levenshtein(ByVal s1 As String, ByVal s2 As String) As Long
'    Dim len1 As Long, len2 As Long
'    Dim i As Long, j As Long
'    Dim cost As Long
'    Dim v0() As Long, v1() As Long
'    Dim temp() As Long
'
'    ' Normalización opcional para apellidos
'    s1 = Trim$(LCase$(s1))
'    s2 = Trim$(LCase$(s2))
'
'    len1 = Len(s1)
'    len2 = Len(s2)
'
'    ' Casos rápidos
'    If len1 = 0 Then
'        Levenshtein = len2
'        Exit Function
'    End If
'
'    If len2 = 0 Then
'        Levenshtein = len1
'        Exit Function
'    End If
'
'    ' Redimensionar matrices
'    ReDim v0(0 To len2)
'    ReDim v1(0 To len2)
'
'    ' Inicializar primera fila
'    For i = 0 To len2
'        v0(i) = i
'    Next i
'
'    ' Bucle principal
'    For i = 1 To len1
'        v1(0) = i
'
'        For j = 1 To len2
'            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
'                cost = 0
'            Else
'                cost = 1
'            End If
'
'            v1(j) = Application.WorksheetFunction.Min( _
'                        v1(j - 1) + 1, _
'                        v0(j) + 1, _
'                        v0(j - 1) + cost)
'        Next j
'
'        ' Intercambiar filas
'        temp = v0
'        v0 = v1
'        v1 = temp
'    Next i
'
'    Levenshtein = v0(len2)
'End Function






'' ------------------------------------------------------------
'' 2. EsVocal: vocal universal (para todos los idiomas)
'' ------------------------------------------------------------
'Public Function EsVocal(ByVal c As String) As Boolean
'    EsVocal = InStr("AEIOUÁÉÍÓÚÀÈÌÒÙÂÊÔÃÕÜ", c) > 0
'End Function


''Public Function BuscarExcepcion(ByVal palabra As String, ByVal idioma As String) As String
''
''    Dim rs As DAO.Recordset
''    Set rs = CurrentDb.OpenRecordset( _
''        "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
''        "WHERE Idioma = '" & idioma & "' AND Palabra = '" & UCase(palabra) & "' AND Activo = True")
''
''    If Not rs.EOF Then
''        BuscarExcepcion = rs!FonemaCompleto
''    Else
''        BuscarExcepcion = ""
''    End If
''
''    rs.Close
''    Set rs = Nothing
''
''End Function
'
'Public Function BuscarExcepcion(ByVal graf As String, ByVal idioma As String) As Byte
'
'    Dim rs As DAO.Recordset
'    Dim sql As String
'
'    sql = "SELECT idFonema FROM tbmDicExcepciones " & _
'          "WHERE Tipo = 'GRAFEMA' " & _
'          "AND Idioma = '" & idioma & "' " & _
'          "AND Grafema = '" & graf & "' " & _
'          "AND Activo = True"
'
'    Set rs = CurrentDb.OpenRecordset(sql)
'
'    If Not rs.EOF Then
'        BuscarExcepcion = rs!idFonema
'    Else
'        BuscarExcepcion = 0
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'End Function
'
'Public Function BuscarExcepcionPalabra(ByVal palabra As String, ByVal idioma As String) As String
'
'    Dim rs As DAO.Recordset
'    Dim sql As String
'
'    sql = "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
'          "WHERE Tipo = 'PALABRA' " & _
'          "AND Idioma = '" & idioma & "' " & _
'          "AND Palabra = '" & UCase(palabra) & "' " & _
'          "AND Activo = True"
'
'    Set rs = CurrentDb.OpenRecordset(sql)
'
'    If Not rs.EOF Then
'        BuscarExcepcionPalabra = rs!FonemaCompleto
'    Else
'        BuscarExcepcionPalabra = ""
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'End Function

'Private Function MF_NormalizarVocales_EN_US(ByVal texto As String) As String
'
'    ' Solo por robustez ante nombres importados
'    texto = Replace(texto, "Á", "A")
'    texto = Replace(texto, "À", "A")
'    texto = Replace(texto, "Ä", "A")
'    texto = Replace(texto, "Â", "A")
'
'    texto = Replace(texto, "É", "E")
'    texto = Replace(texto, "È", "E")
'    texto = Replace(texto, "Ë", "E")
'    texto = Replace(texto, "Ê", "E")
'
'    texto = Replace(texto, "Í", "I")
'    texto = Replace(texto, "Ì", "I")
'    texto = Replace(texto, "Ï", "I")
'    texto = Replace(texto, "Î", "I")
'
'    texto = Replace(texto, "Ó", "O")
'    texto = Replace(texto, "Ò", "O")
'    texto = Replace(texto, "Ö", "O")
'    texto = Replace(texto, "Ô", "O")
'
'    texto = Replace(texto, "Ú", "U")
'    texto = Replace(texto, "Ù", "U")
'    texto = Replace(texto, "Ü", "U")
'    texto = Replace(texto, "Û", "U")
'
'    MF_NormalizarVocales_EN_US = texto
'
'End Function
'
'Private Function MF_NormalizarVocales_EN_US_AF(ByVal texto As String) As String
'
'    ' Solo por robustez ante nombres importados
'    texto = Replace(texto, "Á", "A")
'    texto = Replace(texto, "À", "A")
'    texto = Replace(texto, "Ä", "A")
'    texto = Replace(texto, "Â", "A")
'
'    texto = Replace(texto, "É", "E")
'    texto = Replace(texto, "È", "E")
'    texto = Replace(texto, "Ë", "E")
'    texto = Replace(texto, "Ê", "E")
'
'    texto = Replace(texto, "Í", "I")
'    texto = Replace(texto, "Ì", "I")
'    texto = Replace(texto, "Ï", "I")
'    texto = Replace(texto, "Î", "I")
'
'    texto = Replace(texto, "Ó", "O")
'    texto = Replace(texto, "Ò", "O")
'    texto = Replace(texto, "Ö", "O")
'    texto = Replace(texto, "Ô", "O")
'
'    texto = Replace(texto, "Ú", "U")
'    texto = Replace(texto, "Ù", "U")
'    texto = Replace(texto, "Ü", "U")
'    texto = Replace(texto, "Û", "U")
'
'    MF_NormalizarVocales_EN_US_AF = texto
'
'End Function

'Public Function EsVocal(ByVal c As String) As Boolean
'
''Versión blindada UNICODE
'
'    Dim code As Long
'    code = AscW(c)
'
'    ' Vocales básicas A E I O U
'    If code = &H41 Or code = &H45 Or code = &H49 Or code = &H4F Or code = &H55 Then
'        EsVocal = True: Exit Function
'    End If
'
'    ' Vocales minúsculas (por si acaso)
'    If code = &H61 Or code = &H65 Or code = &H69 Or code = &H6F Or code = &H75 Then
'        EsVocal = True: Exit Function
'    End If
'
'    ' Vocales acentuadas (agudas, graves, circunflejos, diéresis)
'    ' Rango general: U+00C0 – U+00FF (letras latinas extendidas)
'    If code >= &HC0 And code <= &HFF Then
'        Select Case code
'            ' Á É Í Ó Ú
'            Case &HC1, &HC9, &HCD, &HD3, &HDA
'                EsVocal = True: Exit Function
'
'            ' À È Ì Ò Ù
'            Case &HC0, &HC8, &HCC, &HD2, &HD9
'                EsVocal = True: Exit Function
'
'            ' Â Ê Î Ô Û
'            Case &HC2, &HCA, &HCE, &HD4, &HDB
'                EsVocal = True: Exit Function
'
'            ' Ä Ë Ï Ö Ü
'            Case &HC4, &HCB, &HCF, &HD6, &HDC
'                EsVocal = True: Exit Function
'
'            ' Nasales portuguesas: Ã Õ
'            Case &HC3, &HD5
'                EsVocal = True: Exit Function
'        End Select
'    End If
'
'    ' Si no coincide con nada
'    EsVocal = False
'
'End Function

'Public Function EsVocal(ByVal C As String) As Boolean
'
'' Versión no blindada
'
'    Select Case UCase$(C)
'
'        ' Vocales simples
'        Case "A", "E", "I", "O", "U"
'
'        ' Agudas
'        Case "Á", "É", "Í", "Ó", "Ú"
'
'        ' Graves
'        Case "À", "È", "Ì", "Ò", "Ù"
'
'        ' Circunflejos
'        Case "Â", "Ê", "Î", "Ô", "Û"
'
'        ' Diéresis
'        Case "Ä", "Ë", "Ï", "Ö", "Ü"
'
'        ' Nasales portuguesas
'        Case "Ã", "Õ"
'
'            EsVocal = True
'
'        Case Else
'            EsVocal = False
'
'    End Select
'
'End Function

'Public Function EsConsonante(ByVal c As String) As Boolean
'    If c = "" Then
'        EsConsonante = False
'    Else
'        EsConsonante = Not EsVocal(c)
'    End If
'End Function

