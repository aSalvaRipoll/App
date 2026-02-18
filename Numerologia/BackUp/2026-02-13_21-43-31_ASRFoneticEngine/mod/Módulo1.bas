Attribute VB_Name = "Módulo1"

Option Compare Database
Option Explicit

' Caché en memoria para acelerar búsquedas
Private IPA_Cache As Scripting.Dictionary

Public Sub ConstruirCadenaFonemas()

    Dim arrSilabas As Variant
    Dim i As Long
    Dim silaba As String
    Dim ultimaSilaba As String
    Dim SiguienteSilaba As String

    ObjDTO.FonemasFinal = ""

    ' 1. Separar sílabas por "|"
    arrSilabas = Split(ObjDTO.SilabasAcentuadas, "|")

    For i = 0 To UBound(arrSilabas)

        silaba = Trim$(arrSilabas(i))

        ' ---------------------------------------------------------
        ' 2. Detectar separador de palabra (sílaba vacía)
        ' ---------------------------------------------------------
        If silaba = "" Then

            ' Insertar separador estructural de palabra
            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "#"

            ' Ligadura automática (Modo 2)
            If CFG.ModoLigadura = 2 Then

                ' última sílaba real antes del espacio
                ultimaSilaba = BuscarSilabaRealAnterior(arrSilabas, i)

                ' primera sílaba real después del espacio
                SiguienteSilaba = BuscarSilabaRealPosterior(arrSilabas, i)

                If HayLigaduraAutomatica(ultimaSilaba, SiguienteSilaba) Then
                    ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "84,"
                End If

            End If

            ' Pasar a la siguiente sílaba
'            Continue For
            GoTo SiguienteIteracion

        End If

        ' ---------------------------------------------------------
        ' 3. Insertar modificadores prosódicos (acento)
        ' ---------------------------------------------------------
        If InStr(silaba, "(") > 0 Then
            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "80,"   ' acento primario
        ElseIf InStr(silaba, "[") > 0 Then
            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "81,"   ' acento secundario
        Else
            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "82,"   ' átona
        End If

        ' Limpiar marcadores
        silaba = Replace(silaba, "(", "")
        silaba = Replace(silaba, ")", "")
        silaba = Replace(silaba, "[", "")
        silaba = Replace(silaba, "]", "")

        ' ---------------------------------------------------------
        ' 4. Procesar grafemas y ligadura manual
        ' ---------------------------------------------------------
        ProcesarSilaba silaba

        ' ---------------------------------------------------------
        ' 5. Separador silábico
        ' ---------------------------------------------------------
        ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "|"

SiguienteIteracion:
    Next i

    ' ---------------------------------------------------------
    ' 6. Limpieza final
    ' ---------------------------------------------------------
    ObjDTO.FonemasFinal = Replace(ObjDTO.FonemasFinal, "|#", "#")
    ObjDTO.FonemasFinal = Replace(ObjDTO.FonemasFinal, "#|", "#")

    If Right$(ObjDTO.FonemasFinal, 1) = "|" Or Right$(ObjDTO.FonemasFinal, 1) = "#" Then
        ObjDTO.FonemasFinal = Left$(ObjDTO.FonemasFinal, Len(ObjDTO.FonemasFinal) - 1)
    End If

End Sub

'Public Sub ConstruirCadenaFonemas()
'
'    Dim arrPalabras As Variant
'    Dim arrSilabas As Variant
'    Dim palabra As Variant
'    Dim silaba As Variant
'    Dim i As Long
'
'    ObjDTO.FonemasFinal = ""
'
'    ' 1. Separar palabras
'    arrPalabras = Split(ObjDTO.TextoNormalizado, " ")
'
'    For i = 0 To UBound(arrPalabras)
'
'        palabra = arrPalabras(i)
'        arrSilabas = Split(ObjDTO.SilabasAcentuadasPorPalabra(i), "|")
'
'        ' 2. Si no es la primera palabra, evaluar ligadura automática
'        If i > 0 Then
'            If CFG.ModoLigadura = 2 Then
'                If HayLigaduraAutomatica(arrPalabras(i - 1), palabra) Then
'                    ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "84,"
'                End If
'            End If
'        End If
'
'        ' 3. Procesar sílabas de la palabra
'        For Each silaba In arrSilabas
'
'            ' Insertar acento primario/secundario/átona
'            If InStr(silaba, "(") > 0 Then
'                ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "80,"
'            ElseIf InStr(silaba, "[") > 0 Then
'                ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "81,"
'            Else
'                ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "82,"
'            End If
'
'            ' Limpiar marcadores
'            silaba = Replace(silaba, "(", "")
'            silaba = Replace(silaba, ")", "")
'            silaba = Replace(silaba, "[", "")
'            silaba = Replace(silaba, "]", "")
'
'            ' Procesar grafemas y ligadura manual
'            ProcesarSilaba silaba
'
'            ' Separador silábico
'            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "|"
'
'        Next silaba
'
'        ' 4. Separador de palabra
'        ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "#"
'
'    Next i
'
'    ' 5. Limpiar separadores finales
'    ObjDTO.FonemasFinal = Trim$(ObjDTO.FonemasFinal)
'    If Right$(ObjDTO.FonemasFinal, 1) = "#" Then
'        ObjDTO.FonemasFinal = Left$(ObjDTO.FonemasFinal, Len(ObjDTO.FonemasFinal) - 1)
'    End If
'
'End Sub


Function AsignarFonemaBase(grafema As String) As Byte
    Dim rs As DAO.Recordset
    Dim sql As String

    ' Normalizar grafema
    grafema = LCase(grafema)

    ' Consulta para buscar el grafema en la tabla fonética
    sql = "SELECT idFonema FROM Q_FonemasUnion WHERE grafema = '" & grafema & "'"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        ' Encontrado ? devolver ID
        AsignarFonemaBase = rs!idFonema
    Else
        ' No encontrado ? devolver ID especial
        AsignarFonemaBase = 255
    End If

    rs.Close
    Set rs = Nothing
End Function


'---------------------------------------------------------
'                   FUNCIONES AUXILIARES
'---------------------------------------------------------
Private Sub ProcesarSilaba(ByVal silabaCruda As String)

    Dim silaba As String
    Dim ligaduraID As Byte
    Dim i As Long
    Dim grafema As String
    Dim id As Byte

    silaba = Trim$(silabaCruda)

    ' 1. Detectar ligadura manual (solo Modo 1)
    If CFG.ModoLigadura = 1 Then
        ligaduraID = DetectarLigaduraManual(silaba)
        If ligaduraID <> 0 Then
            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & CStr(ligaduraID) & ","
        End If
    End If

    ' 2. Procesar grafemas normalmente
    i = 1
    Do While i <= Len(silaba)

        grafema = DetectarDigrafo(silaba, i)
        id = AsignarFonemaBaseEs(grafema)

        ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & CStr(id) & ","

        i = i + Len(grafema)
    Loop

End Sub

'Private Sub ProcesarSilaba(ByVal silabaCruda As String)
'
'    Dim silaba As String
'    Dim hayLigadura As Boolean
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'
'    silaba = Trim$(silabaCruda)
'
'    ' 1. Detectar ligadura manual (solo Modo 1)
'    If CFG.ModoLigadura = 1 Then
'        hayLigadura = DetectarLigaduraManual(silaba)
'        If hayLigadura Then
'            ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & "84,"   ' ID ligadura
'        End If
'    End If
'
'    ' 2. Procesar grafemas normalmente
'    i = 1
'    Do While i <= Len(silaba)
'
'        grafema = DetectarDigrafo(silaba, i)
'        id = AsignarFonemaBaseEs(grafema)
'
'        ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & CStr(id) & ","
'
'        i = i + Len(grafema)
'    Loop
'
'End Sub

'Sub ProcesarSilaba(silaba As String)
'
'    Dim i As Long
'    Dim grafema As String
'    Dim id As Byte
'
'    i = 1
'    Do While i <= Len(silaba)
'
'        ' 1. Detectar espacio
'        If Mid(silaba, i, 1) = " " Then
'            dto.FonemasFinal = dto.FonemasFinal & "0,"
'            i = i + 1
'            GoTo siguiente
'        End If
'
'        ' 2. Detectar dígrafos (CH, LL, RR, QU, GU…)
'        grafema = DetectarDigrafo(silaba, i)
'
'        ' 3. Obtener ID fonético
'        id = AsignarFonemaBase(grafema)
'
'        ' 4. Añadir a la cadena
'        ObjDTO.FonemasFinal = ObjDTO.FonemasFinal & CStr(id) & ","
'
'        ' 5. Avanzar posición
'siguiente:
'        i = i + Len(grafema)
'
'    Loop
'
'End Sub

Function NormalizarGrafemaEs(ByVal texto As String) As String

    texto = LCase$(texto)

    ' Normalizar acentos
    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ú", "u")

    ' H muda ? eliminar
    texto = Replace(texto, "h", "")

    NormalizarGrafemaEs = texto
End Function

Private Function BuscarSilabaRealAnterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos - 1 To 0 Step -1
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealAnterior = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealAnterior = ""
End Function

Private Function BuscarSilabaRealPosterior(arr As Variant, pos As Long) As String
    Dim j As Long
    For j = pos + 1 To UBound(arr)
        If Trim$(arr(j)) <> "" Then
            BuscarSilabaRealPosterior = Trim$(arr(j))
            Exit Function
        End If
    Next j
    BuscarSilabaRealPosterior = ""
End Function

Function DetectarDigrafoEs(texto As String, pos As Long) As String

    Dim par As String, tri As String, sig As String

    If pos + 1 > Len(texto) Then
        DetectarDigrafoEs = Mid$(texto, pos, 1)
        Exit Function
    End If

    par = Mid$(texto, pos, 2)
    sig = Mid$(texto, pos + 2, 1)

    Select Case par
        Case "ch", "ll", "rr", "qu"
            DetectarDigrafoEs = par
            Exit Function
    End Select

    ' GU + E/I
    If par = "gu" Then
        If sig = "e" Or sig = "i" Then
            DetectarDigrafoEs = "gu"
            Exit Function
        End If
    End If

    ' HI + vocal ? /?/
    If par = "hi" Then
        If sig Like "[aeiou]" Then
            DetectarDigrafoEs = "hi"
            Exit Function
        End If
    End If

    DetectarDigrafoEs = Mid$(texto, pos, 1)
End Function

'Function DetectarDigrafo(texto As String, pos As Long) As String
'
'    Dim par As String
'    par = Mid(texto, pos, 2)
'
'    Select Case par
'        Case "ch", "ll", "rr", "qu"
'            DetectarDigrafo = par
'            Exit Function
'    End Select
'
'    ' GU + E/I
'    If par = "gu" Then
'        Dim siguiente As String
'        siguiente = Mid(texto, pos + 2, 1)
'        If siguiente = "e" Or siguiente = "i" Then
'            DetectarDigrafo = "gu"
'            Exit Function
'        End If
'    End If
'
'    ' Si no es dígrafo ? devolver una sola letra
'    DetectarDigrafo = Mid(texto, pos, 1)
'
'End Function

Function AplicarDecisionesCastellano(ByVal texto As String) As String

    texto = LCase$(texto)

    ' Normalizar acentos
    texto = Replace(texto, "á", "a")
    texto = Replace(texto, "é", "e")
    texto = Replace(texto, "í", "i")
    texto = Replace(texto, "ó", "o")
    texto = Replace(texto, "ú", "u")

    ' -------------------------
    ' TRATAMIENTO DE LA H
    ' -------------------------
    Select Case CFG.ModoH

        Case 0, 1 '"H muda siempre", "Ninguna"
            texto = Replace(texto, "h", "")

        Case 2 '"H en préstamos"
'            texto = ProcesarHPrestamos(texto)

        Case 3 '"H aspirada"
'            texto = ReemplazarHAspirada(texto)

        Case Else
        
    End Select

    AplicarDecisionesCastellano = texto
End Function


Function AsignarFonemaBaseEs(grafema As String, _
                             Optional ByVal sig As String = "", _
                             Optional ByVal ant As String = "") As Byte

    grafema = LCase$(grafema)

    ' VOCAL
    Select Case grafema
        Case "a": AsignarFonemaBaseEs = 1: Exit Function
        Case "e": AsignarFonemaBaseEs = 2: Exit Function
        Case "i": AsignarFonemaBaseEs = 4: Exit Function
        Case "o": AsignarFonemaBaseEs = 5: Exit Function
        Case "u": AsignarFonemaBaseEs = 7: Exit Function
    End Select

    ' OCLUSIVAS
    Select Case grafema
        Case "p": AsignarFonemaBaseEs = 30: Exit Function
        Case "b", "v": AsignarFonemaBaseEs = 31: Exit Function
        Case "t": AsignarFonemaBaseEs = 32: Exit Function
        Case "d": AsignarFonemaBaseEs = 33: Exit Function
    End Select

    ' K
    If grafema = "k" Then
        AsignarFonemaBaseEs = 34
        Exit Function
    End If

    ' C
    If grafema = "c" Then
        If sig = "e" Or sig = "i" Then
            AsignarFonemaBaseEs = 46   ' /?/
        Else
            AsignarFonemaBaseEs = 34   ' /k/
        End If
        Exit Function
    End If

    ' QU
    If grafema = "qu" Then
        AsignarFonemaBaseEs = 34
        Exit Function
    End If

    ' G
    If grafema = "g" Then
        If sig = "e" Or sig = "i" Then
            AsignarFonemaBaseEs = 48   ' /x/
        Else
            AsignarFonemaBaseEs = 35   ' /g/
        End If
        Exit Function
    End If

    ' GU
    If grafema = "gu" Then
        AsignarFonemaBaseEs = 35
        Exit Function
    End If

    ' FRICATIVAS
    Select Case grafema
        Case "f": AsignarFonemaBaseEs = 40: Exit Function
        Case "z": AsignarFonemaBaseEs = 46: Exit Function
        Case "s": AsignarFonemaBaseEs = 42: Exit Function
        Case "j": AsignarFonemaBaseEs = 48: Exit Function
    End Select

    ' NASALES
    Select Case grafema
        Case "m": AsignarFonemaBaseEs = 36: Exit Function
        Case "n": AsignarFonemaBaseEs = 37: Exit Function
        Case "ñ": AsignarFonemaBaseEs = 38: Exit Function
    End Select

    ' LATERALES
    If grafema = "l" Then AsignarFonemaBaseEs = 62: Exit Function
    If grafema = "ll" Then AsignarFonemaBaseEs = 63: Exit Function

    ' VIBRANTES
    If grafema = "rr" Then
        AsignarFonemaBaseEs = 60
        Exit Function
    End If

    If grafema = "r" Then
        If ant = "" Or Not ant Like "[aeiou]" Then
            AsignarFonemaBaseEs = 60   ' inicio de palabra
        Else
            AsignarFonemaBaseEs = 59   ' intervocálica
        End If
        Exit Function
    End If

    ' AFRICADA
    If grafema = "ch" Then
        AsignarFonemaBaseEs = 57
        Exit Function
    End If

    ' Y /?/
    If grafema = "y" Or grafema = "hi" Then
        AsignarFonemaBaseEs = 45
        Exit Function
    End If


Select Case CFG.ModoSibilantes

    Case 0, 1 '"Distinción"
        ' c/z ? ?, s ? s
        If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46
        If grafema = "z" Then AsignarFonemaBaseEs = 46
        If grafema = "s" Then AsignarFonemaBaseEs = 42

    Case 2 '"Seseo"
        ' c/z/s ? s
        If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 42
        If grafema = "z" Then AsignarFonemaBaseEs = 42
        If grafema = "s" Then AsignarFonemaBaseEs = 42

    Case 3 '"Ceceo"
        ' c/z/s ? ?
        If grafema = "c" And (sig = "e" Or sig = "i") Then AsignarFonemaBaseEs = 46
        If grafema = "z" Then AsignarFonemaBaseEs = 46
        If grafema = "s" Then AsignarFonemaBaseEs = 46

    Case Else
    
End Select

Select Case CFG.ModoLateral

    Case 1 '"Lleísmo"
        If grafema = "ll" Then AsignarFonemaBaseEs = 63   ' /?/
        If grafema = "y" Then AsignarFonemaBaseEs = 45    ' /?/

    Case 2, "Yeísmo"
        If grafema = "ll" Then AsignarFonemaBaseEs = 45   ' /?/
        If grafema = "y" Then AsignarFonemaBaseEs = 45    ' /?/

    Case 3 '"Yeísmo rehilado"
        If grafema = "ll" Then AsignarFonemaBaseEs = 72   ' /?/
        If grafema = "y" Then AsignarFonemaBaseEs = 72    ' /?/

    Case Else
    
End Select

Select Case CFG.ModoX

    Case "ks"
        If grafema = "x" Then AsignarFonemaBaseEs = 34  ' /ks/ ? no monofonémico, pero puedes mapearlo

    Case "s"
        If grafema = "x" Then AsignarFonemaBaseEs = 42

    Case "?"
        If grafema = "x" Then AsignarFonemaBaseEs = 44

    Case "x"
        If grafema = "x" Then AsignarFonemaBaseEs = 48

    Case "contextual"
        ' Aquí podemos hacer magia más adelante
        ' México ? /x/
        ' examen ? /ks/
        ' exacto ? /gz/
        ' xilófono ? /s/
    Case Else
    
End Select
'
'    Select Case CFG.ModoLigadura
'
'        Case 0
'            ' Sin ligaduras, ignorar cualquier intento de insertarlas
'
'        Case 1
'            ' Ligaduras solo manuales (modo simple)
'            ' Si el usuario marca una ligadura, se respeta
'            ' Si no, no se añade ninguna
'
'        Case 2
'            ' Ligaduras automáticas y manuales (modo simple)
'            ' 2 = automáticas básicas
'
'        Case Else
'            ' Futuras expansiones
'            ' 3 = automáticas avanzadas
'
'    End Select


    ' Si no se reconoce
    AsignarFonemaBaseEs = 255
End Function

Private Function DetectarLigaduraManual(ByRef silaba As String) As Byte

    DetectarLigaduraManual = 0
    
    If InStr(silaba, "_") > 0 Then
        silaba = Replace(silaba, "_", "")
        DetectarLigaduraManual = 84   ' ID de ligadura
    End If

End Function

Private Function HayLigaduraAutomatica(ultimaSilaba As String, primeraSilaba As String) As Boolean

    Dim ult As String
    Dim pri As String

    HayLigaduraAutomatica = False

    If ultimaSilaba = "" Or primeraSilaba = "" Then Exit Function

    ult = Right$(Trim$(ultimaSilaba), 1)
    pri = Left$(Trim$(primeraSilaba), 1)

    ' Vocal + vocal
    If ult Like "[aeiouáéíóúü]" And pri Like "[aeiouáéíóúü]" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

    ' Vocal + h + vocal
    If ult Like "[aeiouáéíóúü]" And primeraSilaba Like "h[aeiouáéíóúü]*" Then
        HayLigaduraAutomatica = True
        Exit Function
    End If

End Function

'Private Function HayLigaduraAutomatica(palabraAnt As String, palabraSig As String) As Boolean
'
'    Dim ult As String
'    Dim pri As String
'
'    HayLigaduraAutomatica = False
'
'    If palabraAnt = "" Or palabraSig = "" Then Exit Function
'
'    ult = Right$(palabraAnt, 1)
'    pri = Left$(palabraSig, 1)
'
'    ' Vocal + vocal
'    If ult Like "[aeiouáéíóúü]" And pri Like "[aeiouáéíóúü]" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'    ' Vocal + h + vocal
'    If ult Like "[aeiouáéíóúü]" And Left$(palabraSig, 2) Like "h[aeiouáéíóúü]" Then
'        HayLigaduraAutomatica = True
'        Exit Function
'    End If
'
'End Function

Function DetectarDigrafo(texto As String, pos As Long) As String

    Dim arr As Variant
    Dim d As Variant
    Dim largo As Long
    Dim candidato As String

    arr = ObtenerArrayDigrafos()

    ' Probar primero los más largos
    For Each d In arr
        largo = Len(d)
        If pos + largo - 1 <= Len(texto) Then
            candidato = Mid$(texto, pos, largo)
            If candidato = d Then
                DetectarDigrafo = d
                Exit Function
            End If
        End If
    Next d

    ' Si no coincide ningún dígrafo ? letra suelta
    DetectarDigrafo = Mid$(texto, pos, 1)

End Function

Private Sub CargaCombos(frm As Form)

    With frm

        ' Limpiar combos
        .cmbVocales.Value = ""
        .cmbTrataH.Value = ""
        .cmbGenero.Value = ""
        .cmbNumero.Value = ""
        ' ... los que correspondan

        ' Cargar valores según idioma
        ' Ejemplo para Castellano
        .cmbVocales.Clear
        .cmbVocales.AddItem "0;Ninguna"
        .cmbVocales.AddItem "1;Acentos Castellanos"
        .cmbVocales.AddItem "2;Acentos Graves y Agudos"
        .cmbVocales.AddItem "3;Todos los acentos y marcadores"

        ' ... resto de combos según idioma

        ' Cargar el ListBox con SQL dinámico
        Dim strSQL As String
        strSQL = CreaSQL("Castellano")
        .lstDigrafos.RowSource = strSQL

    End With

End Sub

Public Function GenerarIPA() As String

    Dim salida As String
    Dim arrPalabras As Variant
    Dim arrSilabas As Variant
    Dim arrFonemas As Variant
    Dim palabra As Variant
    Dim silaba As Variant
    Dim f As Variant
    Dim id As Long
    Dim ipa As String

    salida = ""

    ' 1. Separar palabras por "#"
    arrPalabras = Split(ObjDTO.FonemasFinal, "#")

    For Each palabra In arrPalabras

        If Trim$(palabra) = "" Then GoTo SiguientePalabra

        ' 2. Separar sílabas por "|"
        arrSilabas = Split(palabra, "|")

        For Each silaba In arrSilabas

            If Trim$(silaba) = "" Then GoTo SiguienteSilaba

            ' 3. Separar fonemas/modificadores por ","
            arrFonemas = Split(Trim$(silaba), ",")

            ipa = ""

            For Each f In arrFonemas

                If Trim$(f) = "" Then GoTo SiguienteFonema

                If IsNumeric(f) Then
                    id = CLng(f)

                    Select Case id

                        ' ------------------------------
                        ' MODIFICADORES PROSÓDICOS
                        ' ------------------------------
                        Case 80: ipa = ipa & "'"     ' acento primario
                        Case 81: ipa = ipa & "?"     ' acento secundario
                        Case 82:                     ' átona ? no se muestra
                        Case 84: ipa = ipa & "?"     ' ligadura

                        ' ------------------------------
                        ' FONEMAS REALES
                        ' ------------------------------
                        Case Else
                            ipa = ipa & ObtenerIPA(id)

                    End Select

                End If

SiguienteFonema:
            Next f

            ' 4. Añadir punto silábico
            salida = salida & ipa & "."

SiguienteSilaba:
        Next silaba

        ' 5. Espacio entre palabras
        salida = salida & " "

SiguientePalabra:
    Next palabra

    ' 6. Limpieza final
    salida = Trim$(salida)

    ' Quitar punto final si existe
    If Right$(salida, 1) = "." Then
        salida = Left$(salida, Len(salida) - 1)
    End If

    GenerarIPA = salida

End Function


Public Function ObtenerIPA(ByVal idFonema As Long) As String
    Dim rs As DAO.Recordset

    ' Inicializar caché si es la primera vez
    If IPA_Cache Is Nothing Then
        Set IPA_Cache = CreateObject("Scripting.Dictionary")
    End If

    ' Si ya está en caché ? devolverlo directamente
    If IPA_Cache.Exists(idFonema) Then
        ObtenerIPA = IPA_Cache(idFonema)
        Exit Function
    End If

    ' ID desconocido o especial
    If idFonema = 255 Then
        IPA_Cache.Add idFonema, ""   ' vacío
        ObtenerIPA = ""
        Exit Function
    End If

    ' Buscar en la tabla qryFonemasValor
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT IPA FROM qryFonemasValor WHERE ID=" & idFonema & ";", _
        dbOpenSnapshot)

    If Not (rs.EOF And rs.BOF) Then
        IPA_Cache.Add idFonema, Nz(rs!ipa, "")
        ObtenerIPA = Nz(rs!ipa, "")
    Else
        ' Si no existe el ID ? devolver vacío
        IPA_Cache.Add idFonema, ""
        ObtenerIPA = ""
    End If

    rs.Close
    Set rs = Nothing
End Function

Public Function ObtenerArrayDigrafos() As Variant
'    If CFG.ListaDigrafos = "" Then
'        ObtenerArrayDigrafos = Array()
'    Else
'        ObtenerArrayDigrafos = Split(CFG.ListaDigrafos, ";")
'    End If
End Function


'Function DetectarDigrafo(texto As String, pos As Long) As String
'
'    Dim d As Variant
'    Dim largo As Long
'    Dim candidato As String
'
'    ' Probar primero los más largos (trígrafos)
'    For Each d In DigrafosActivos
'        largo = Len(d)
'        If pos + largo - 1 <= Len(texto) Then
'            candidato = Mid$(texto, pos, largo)
'            If candidato = d Then
'                DetectarDigrafo = d
'                Exit Function
'            End If
'        End If
'    Next d
'
'    ' Si no coincide ningún dígrafo ? letra suelta
'    DetectarDigrafo = Mid$(texto, pos, 1)
'
'End Function



