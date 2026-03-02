Attribute VB_Name = "bas_Motor_ES_Main"

Option Compare Database
Option Explicit

' ============================================================
'   MOTOR DE SILABEO 4.1 — ORTOGRÁFICO + MORFOSILÁBICO + ACENTUACIÓN
'   VERSIÓN CASTELLANO (ESPAÑOL ESTANDARD)
'
'   Autor: Alba Salvá Ripoll
'
'   Este módulo implementa:
'   ? Silabeo ortográfico completo (RAE 2010)
'   ? Silabeo morfosilábico (prefijos)
'   ? Diptongos, triptongos, hiatos con tilde
'   ? Dígrafos indivisibles: ch, ll, rr
'   ? Casos especiales: qu, gu, güe/güi, tl
'   ? Sílaba tónica
'   ? Tracking de depuración (DebugMotor)
'
'   Variables de control:
'       usarSilabeoMorfologico   ? True = usar prefijos
'       modoPrefijosEstrictos    ? True = usar lista interna completa
'       respetarPrefijos         ? True = separar prefijo como sílaba
'       DebugMotor                ? True = activar logs
'
' ============================================================

Private usarSilabeoMorfologico As Boolean
Private modoPrefijosEstrictos As Boolean
Private respetarPrefijos As Boolean

' Prefijos personalizados (modo flexible)
Private prefijosPersonalizados As Variant

' Prefijos canónicos (modo estricto)
Private prefijosEstrictos As Variant

Private prefijosCargados As Boolean

Private Const strSQL As String = "SELECT Prefijo FROM qryPrefijos " & _
          "WHERE Activo = 1 " & _
            "AND Tipo Like 'auténtico' " & _
            "AND [ES] = true " & _
          "ORDER BY Len(Prefijo) DESC, Prefijo ASC"

' ============================================================
'   ENTRADA PRINCIPAL DEL MOTOR (ESPAÑOL)
' ============================================================
' ----------------------------------------------------------------
' Procedimiento: Entrada_Motor_GL
' Propósito:     Punto de entrada al motor fonético del español general
' Tipo proc.:    Function
' Acceso proc.:  Public

' Parameter Texto (String): Texto que se recibe (nombre o apellido español general)

' Tipo retorno: String -> Texto que contiene la lista de fonemas
'   resultado de la conversión

' Autor:        Alba Salvá
' Fecha:        11/02/2026
' ----------------------------------------------------------------

Public Function Entrada_Motor_ES(texto As String) As String

    Set ObjDTO = New clsDTO_Motor

    ' 1) Asignamos el texto recibido y
    '    Normalización (dentro del DTO)
    ObjDTO.TextoOriginal = texto
    ObjDTO.NormalizaEntrada

    ' 2) Silabeo automático
    Call SilabearFrase

    ' 3) Detectar tónicas
    Call CalcularTonicas

    ' 4) Detectar secundarias
    Call CalcularSecundarias

    ' 5) Marcar Tónicas y Secundarias
    Call MarcarTonicaYSecundariaEnCadena

    ObjDTO.SilabasFinal = ObjDTO.SilabasAcentuadas

    ' 6) Generar fonética
    Call ConstruirCadenaFonemas_ES

'    '    << Eliminar en producción >>
'    Call MF_DebugDTO("Silabear")

    ' 7) Devolver resultado (texto plano)
    '    << Eliminar en producción >>
    Entrada_Motor_ES = ObjDTO.SilabasAuto

End Function

'Private Function SilabearFrase(ByVal frase As String) As String
Private Sub SilabearFrase()
    
    Dim frase As String
    Dim palabras() As String
    Dim resultado As String
    Dim i As Long
    Dim limpia As String
    Dim sil As String

    ' Tratamiento prefijos
'       usarSilabeoMorfologico   => True = usar prefijos
    usarSilabeoMorfologico = True

'       modoPrefijosEstrictos    => True = usar lista interna completa
    modoPrefijosEstrictos = True
'       respetarPrefijos         => True = separar prefijo como sílaba
    respetarPrefijos = True

    frase = ObjDTO.TextoNormalizado

    If DebugMotor Then
'        InitLog
        addLog
        addLog "============================================================="
        addLog "        LOG DE DEPURACIÓN SILABEO 5.0"
        addLog "============================================================="
        addLog "Entrada: '" & frase & "'"
'        addLog "Normalizado: '" & T & "'"
        addLog "Longitud: " & Len(frase) & " letras y espacios."
    End If
    
    palabras = Split(frase, " ")

    For i = LBound(palabras) To UBound(palabras)
        'limpia = LimpiarPalabra(palabras(i))
        limpia = Trim$(palabras(i))
        
        If limpia <> "" Then
            sil = SilabearPalabra(limpia)

            ' Añadir separador de palabra:  |   |
            If resultado = "" Then
                resultado = sil
            Else
                resultado = resultado & " |   | " & sil
            End If
        End If
    Next i

    'SilabearFrase = resultado
    ObjDTO.SilabasAuto = resultado
    
    If DebugMotor Then
        addLog
        addLog "Resultado final: " & ObjDTO.SilabasAuto
        addLog "============================================================="
        addLog "                   FIN LOG DE DEPURACIÓN"
        addLog "============================================================="
'        PrintLog
    End If
End Sub

Private Function SilabearPalabra(ByVal texto As String) As String
    Dim t As String
    
    t = LCase$(Trim$(texto))

    If DebugMotor Then
'        InitLog
        addLog
        addLog "============================================================="
'        addLog "        LOG DE DEPURACIÓN SILABEO 4.1"
        addLog "  SÍLABAS DE '" & texto & "'"
        addLog "============================================================="
        addLog "Entrada: '" & texto & "'"
        addLog "Normalizado: '" & t & "'"
        addLog "Longitud: " & Len(t) & " letras."
    End If

    If usarSilabeoMorfologico Then
        SilabearPalabra = SilabearMorfologico(t)
    Else
        SilabearPalabra = SilabearOrtog(t)
    End If

    If DebugMotor Then
        addLog
        addLog "Resultado palabra: " & SilabearPalabra
        addLog "============================================================="
        addLog " FIN SILABEO '" & texto & "'"
        addLog "============================================================="
        'PrintLog
    End If
End Function

Private Function SilabearOrtog(ByVal t As String) As String
    Dim nucIni() As Byte, nucFin() As Byte
    Dim silIni() As Byte, silFin() As Byte
    Dim nNuc As Byte, i As Byte
    Dim silabas() As String

    If t = "y" Then
        SilabearOrtog = "y"
        Exit Function
    End If

    LocalizarNucleosOrtog t, nucIni, nucFin, nNuc

    ReDim silIni(1 To nNuc)
    ReDim silFin(1 To nNuc)

    CalcularSilabas t, nucIni, nucFin, nNuc, silIni, silFin

    ReDim silabas(1 To nNuc)
    For i = 1 To nNuc
        silabas(i) = Mid$(t, silIni(i), silFin(i) - silIni(i) + 1)
        If DebugMotor Then addLog "Sílaba " & i & ": " & silabas(i)
    Next i

    SilabearOrtog = Join(silabas, " | ")
End Function

Private Function SilabearMorfologico(ByVal t As String) As String
' Prefijos
    Dim pref As String
    Dim resto As String

    If Not respetarPrefijos Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    pref = DetectarPrefijo_ES(t)

    If pref = "" Then
        SilabearMorfologico = SilabearOrtog(t)
        Exit Function
    End If

    resto = Mid$(t, Len(pref) + 1)

    If DebugMotor Then
        addLog "Prefijo detectado: " & pref
        addLog "Resto: " & resto
    End If

    SilabearMorfologico = pref & " | " & SilabearOrtog(resto)
End Function

Private Function DetectarPrefijo_ES(ByVal t As String) As String
    Dim p As Variant
    Dim resto As String
    Dim primera As String, segunda As String

    If Not prefijosCargados Then CargarPrefijos

    For Each p In prefijosEstrictos

        ' 1) El prefijo debe coincidir con el inicio
        If Left$(t, Len(p)) <> p Then GoTo SiguienteP

        ' 2) Obtener el resto de la palabra
        resto = Mid$(t, Len(p) + 1)
        If Len(resto) = 0 Then GoTo SiguienteP

        primera = Left$(resto, 1)
        segunda = Mid$(resto, 2, 1)

        ' ============================================================
        '   CASO ESPECIAL: PREFIJO "a-" (privativo)
        ' ============================================================
        If p = "a" Then

            ' 2.1) La base debe empezar por vocal
            If Not EsVocal(primera) Then
                GoTo SiguienteP
            End If

            ' 2.2) No puede romper un diptongo
            If EsDiptongo(Left$(t, 1), Mid$(t, 2, 1)) Then
                GoTo SiguienteP
            End If

            ' 2.3) El ataque resultante debe ser válido
            If Not AtaqueSilabicoValido(primera, segunda) Then
                GoTo SiguienteP
            End If
            
            DetectarPrefijo_ES = p
            Exit Function
        End If

        ' ============================================================
        '   PREFIJOS NORMALES (anti-, inter-, sub-, re-, etc.)
        ' ============================================================

        ' 3) El ataque resultante debe ser válido
        If Not AtaqueSilabicoValido(primera, segunda) Then
            GoTo SiguienteP
        End If

        DetectarPrefijo_ES = p
        Exit Function

SiguienteP:
    Next p

    DetectarPrefijo_ES = ""
End Function

Private Function AtaqueSilabicoValido(ByVal a As String, ByVal b As String) As Boolean
    Dim grupo As String
    grupo = LCase$(a & b)

    Select Case grupo
        Case "pr", "pl", "br", "bl", "tr", "dr", "cr", "cl", "gr", "gl", "fr", "fl"
            AtaqueSilabicoValido = True

        ' Opcional: permitir "tl" en préstamos cultos
        Case "tl"
            AtaqueSilabicoValido = True

        Case Else
            AtaqueSilabicoValido = False
    End Select
End Function

'Private Function DetectarPrefijo(ByVal t As String) As String
'    Dim p As Variant
'
'    If Not prefijosCargados Then CargarPrefijos
'
'    For Each p In prefijosEstrictos
'
'        ' Si la palabra es exactamente igual al prefijo ? NO separar
'        If Len(t) = Len(p) Then
''            GoTo SiguientePrefijo
'            Exit For
'        End If
'
'        If Left$(t, Len(p)) = p Then
'            DetectarPrefijo = p
'            Exit Function
'        End If
'
'SiguientePrefijo:
'    Next p
'
'    DetectarPrefijo = ""
'End Function

Private Sub CargarPrefijos()
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim i As Long

    If prefijosCargados Then Exit Sub

    sql = strSQL
    
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        rs.MoveLast
        ReDim prefijosEstrictos(1 To rs.RecordCount)
        rs.MoveFirst

        i = 1
        Do Until rs.EOF
            prefijosEstrictos(i) = LCase$(rs!Prefijo)
            i = i + 1
            rs.MoveNext
        Loop
    End If

    rs.Close
    prefijosCargados = True
End Sub

Private Sub LocalizarNucleosOrtog(ByVal t As String, _
                                  ByRef nucIni() As Byte, _
                                  ByRef nucFin() As Byte, _
                                  ByRef nNuc As Byte)

    Dim i As Byte, L As Byte
    Dim c1 As String, c2 As String, c3 As String

    L = Len(t)
    ReDim nucIni(1 To L)
    ReDim nucFin(1 To L)
    nNuc = 0

    If DebugMotor Then
        addLog
        addLog "---------------------------------------"
        addLog " Procedimiento LocalizarNucleosOrtog"
    End If

    i = 1
    Do While i <= L
        c1 = Mid$(t, i, 1)

        If EsVocal(c1) Then

            ' Intentar triptongo
            If i + 2 <= L Then
                c2 = Mid$(t, i + 1, 1)
                c3 = Mid$(t, i + 2, 1)
                If EsTriptongo(c1, c2, c3) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 2
                    If DebugMotor Then addLog "Triptongo: " & c1 & c2 & c3
                    i = i + 3
                    GoTo Siguiente
                End If
            End If

            ' Intentar diptongo
            If i + 1 <= L Then
                c2 = Mid$(t, i + 1, 1)
                If EsDiptongo(c1, c2) Then
                    nNuc = nNuc + 1
                    nucIni(nNuc) = i
                    nucFin(nNuc) = i + 1
                    If DebugMotor Then addLog "Diptongo: " & c1 & c2
                    i = i + 2
                    GoTo Siguiente
                End If
            End If

            ' Vocal sola
            nNuc = nNuc + 1
            nucIni(nNuc) = i
            nucFin(nNuc) = i
            If DebugMotor Then addLog "Vocal sola: " & c1
            i = i + 1

        Else
            i = i + 1
        End If

Siguiente:
    Loop

    If DebugMotor Then
        addLog "Total núcleos: " & nNuc
        addLog " Fin LocalizarNucleosOrtog"
        addLog "---------------------------------------"
    End If
End Sub

Private Sub CalcularSilabas(ByVal t As String, _
                            ByRef nucIni() As Byte, _
                            ByRef nucFin() As Byte, _
                            ByVal nNuc As Byte, _
                            ByRef silIni() As Byte, _
                            ByRef silFin() As Byte)

' Con control de dígrafos

    Dim i As Byte, L As Byte
    Dim a As Byte, b As Byte
    Dim k As Byte
    Dim c1 As String, c2 As String, grupo As String

    L = Len(t)
    silIni(1) = 1

    For i = 1 To nNuc - 1
        a = nucFin(i)
        b = nucIni(i + 1)

        k = IIf(b > a + 1, b - a - 1, 0)

        If DebugMotor Then
            addLog
            addLog "---- Frontera entre núcleo " & i & " y " & (i + 1)
            addLog "Consonantes entre medias: " & k
        End If

        Select Case k

            Case 0
                silFin(i) = a
                silIni(i + 1) = a + 1

            Case 1
                silFin(i) = a
                silIni(i + 1) = a + 1

            Case 2
                c1 = Mid$(t, a + 1, 1)
                c2 = Mid$(t, a + 2, 1)
                grupo = c1 & c2

                ' Dígrafos indivisibles
                If grupo = "rr" Or grupo = "ll" Or grupo = "ch" Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                    GoTo Siguiente
                End If

                ' tl ? siempre se separa
                If grupo = "tl" Then
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                    GoTo Siguiente
                End If

                ' Grupos de ataque válidos
                If EsGrupoAtaque(grupo) Then
                    silFin(i) = a
                    silIni(i + 1) = a + 1
                Else
                    silFin(i) = a + 1
                    silIni(i + 1) = a + 2
                End If

            Case 3
                silFin(i) = a + 1
                silIni(i + 1) = a + 2

            Case Else
                silFin(i) = a + 2
                silIni(i + 1) = a + 3

        End Select

Siguiente:
    Next i

    silFin(nNuc) = L
End Sub

Private Function EsVocal(ByVal c As String) As Boolean
    EsVocal = InStr("aeiouáéíóúü", c) > 0
End Function

Private Function EsVocalFuerte(ByVal c As String) As Boolean
    EsVocalFuerte = InStr("aáeéoó", c) > 0
End Function

Private Function EsVocalDebil(ByVal c As String) As Boolean
    EsVocalDebil = InStr("iíuúü", c) > 0
End Function

Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
    EsVocalDebilTonica = InStr("íú", c) > 0
End Function

Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
    ' Si alguna no es vocal ? no hay diptongo
    If Not EsVocal(v1) Or Not EsVocal(v2) Then Exit Function

    ' Dos vocales fuertes ? nunca diptongo
    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then Exit Function

    ' *** REGLA ESPECIAL ***
    ' Dos vocales débiles (i, u, ü) ? SIEMPRE diptongo,
    ' incluso si una lleva tilde (caso güí de "lingüística")
    If EsVocalDebil(v1) And EsVocalDebil(v2) Then
        EsDiptongo = True
        Exit Function
    End If

    ' Si una es débil tónica y la otra es fuerte ? hiato
    If (EsVocalDebilTonica(v1) And EsVocalFuerte(v2)) _
    Or (EsVocalFuerte(v1) And EsVocalDebilTonica(v2)) Then
        Exit Function
    End If

    ' En cualquier otro caso ? diptongo
    EsDiptongo = True
End Function

Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
       And EsVocalFuerte(v2) _
       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
        EsTriptongo = True
    End If
End Function

Private Function EsGrupoAtaque(ByVal g As String) As Boolean
    Dim AC As Variant
    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
               "pl", "bl", "cl", "gl", "fl")
    EsGrupoAtaque = (UBound(Filter(AC, g)) >= 0)
End Function

Private Function SilabaTonica(ByVal palabra As String) As Byte
    Dim sil As String
    Dim silabas() As String
    Dim i As Byte

    sil = SilabearPalabra(palabra)
    silabas = Split(sil, " | ")

    For i = 0 To UBound(silabas)
        If tieneTilde(silabas(i)) Then
            SilabaTonica = i + 1
            Exit Function
        End If
    Next i

    If TerminaEnVocalNS(palabra) Then
        SilabaTonica = UBound(silabas)
    Else
        SilabaTonica = UBound(silabas) + 1
    End If
End Function

Private Function tieneTilde(ByVal silaba As String) As Boolean
    tieneTilde = (InStr(silaba, "á") > 0 Or _
                  InStr(silaba, "é") > 0 Or _
                  InStr(silaba, "í") > 0 Or _
                  InStr(silaba, "ó") > 0 Or _
                  InStr(silaba, "ú") > 0)
End Function

Private Function TerminaEnVocalNS(ByVal s As String) As Boolean
    Dim c As String
    c = Right$(s, 1)
    TerminaEnVocalNS = (EsVocal(c) Or c = "n" Or c = "s")
End Function


'=================================================================
'=================================================================
'                 SECCIÓN MÓDULO ACENTOS
'=================================================================
'=================================================================

' ============================================================
'   DETECTAR SÍLABAS TÓNICAS
' ============================================================
Private Sub CalcularTonicas()

    Dim tGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then
            ' palabra real
            Dim w As Collection
            Set w = elementos(i)

            Dim tLocal As Long
            tLocal = DetectarTonica(w)

            If tLocal > 0 Then
                tGlobal.Add globalIndex + tLocal
            End If

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasTonicas = JoinCollection(tGlobal)

End Sub


' ============================================================
'   DETECTAR SÍLABAS SECUNDARIAS (pueden ser varias)
' ============================================================
Private Sub CalcularSecundarias()

    Dim sGlobal As New Collection
    Dim elementos As Collection
    Set elementos = ObtenerPalabrasDesdeSilabasAuto()

    Dim globalIndex As Byte
    Dim i As Byte
    Dim tLocal As Byte

    globalIndex = 0

    For i = 1 To elementos.count

        If TypeName(elementos(i)) = "Collection" Then

            Dim w As Collection
            Set w = elementos(i)

            ' Detectar tónica local
            tLocal = DetectarTonica(w)

            ' Detectar secundarias locales
            Dim secs As Collection
            Set secs = DetectarSecundarias(w, tLocal)

            ' Convertir secundarias locales a globales
            Dim x As Variant
            For Each x In secs
                sGlobal.Add globalIndex + CByte(x)
            Next x

            globalIndex = globalIndex + w.count

        Else
            ' HUECO
            globalIndex = globalIndex + 1
        End If

    Next i

    ObjDTO.SilabasSecundarias = JoinCollection(sGlobal)

End Sub


' ============================================================
'   MARCAR TÓNICAS Y SECUNDARIAS EN LA CADENA FINAL
' ============================================================
Private Sub MarcarTonicaYSecundariaEnCadena()

    Dim sils As Variant
    Dim i As Byte
    Dim out() As String

    Dim t As Variant
    Dim x As Variant

    sils = Split(ObjDTO.SilabasAuto, " | ")

    ReDim out(LBound(sils) To UBound(sils))

    For i = LBound(sils) To UBound(sils)
        out(i) = sils(i)   ' copia sin marcar
    Next i

    ' 1) TÓNICAS
    If ObjDTO.SilabasTonicas <> "" Then

        t = Split(ObjDTO.SilabasTonicas, ",")

        For Each x In t
            Dim idx As Long
            idx = CByte(x) - 1   ' arrays base 0

            If idx >= LBound(out) And idx <= UBound(out) Then
                If Trim$(out(idx)) <> "" Then
                    out(idx) = "( " & out(idx) & " )"
                End If
            End If
        Next x
    End If

    ' 2) SECUNDARIAS
    If ObjDTO.SilabasSecundarias <> "" Then

        Dim s As Variant
        Dim y As Variant
        Dim idx2 As Byte

        s = Split(ObjDTO.SilabasSecundarias, ",")

        For Each y In s

            idx2 = CByte(y) - 1

            If idx2 >= LBound(out) And idx2 <= UBound(out) Then
                If Trim$(out(idx2)) <> "" Then
                    out(idx2) = "[ " & out(idx2) & " ]"
                End If
            End If
        Next y
    End If

    ' 3) UNIR RESULTADO
    ObjDTO.SilabasAcentuadas = Join(out, " | ")

End Sub

'-------------------------------------------------------------
'             AUXILIARES TÓNICAS Y SECUNDARIAS
'-------------------------------------------------------------
Private Function DetectarTonica(w As Collection) As Byte

    Dim i As Byte
    Dim ultima As String

    For i = 1 To w.count
        If tieneTilde(w(i)) Then
            DetectarTonica = i
            Exit Function
        End If
    Next i

    ultima = Right$(w(w.count), 1)

    If ultima Like "[aeiouns]" Then
        DetectarTonica = w.count - 1
    Else
        DetectarTonica = w.count
    End If

End Function

Private Function DetectarSecundarias(w As Collection, tPos As Byte) As Collection

    Dim secs As New Collection
    Dim n As Byte
    Dim pos2 As Byte

    n = w.count

    ' Palabras de 1–3 sílabas ? sin secundaria
    If n < 4 Then
        Set DetectarSecundarias = secs
        Exit Function
    End If

    ' Primera secundaria SIEMPRE en la sílaba 1
    secs.Add 1

    ' Palabras de 6+ sílabas ? segunda secundaria
    If n >= 6 Then
        pos2 = tPos - 2   ' dos antes de la tónica

        If pos2 > 1 Then
            secs.Add pos2
        End If
    End If

    Set DetectarSecundarias = secs

End Function


' ============================================================
'   OBTENER PALABRAS DESDE SILABAS AUTO
' ============================================================
Private Function ObtenerPalabrasDesdeSilabasAuto() As Collection

    Dim resultado As New Collection
    Dim palabraActual As New Collection

    Dim sils As Variant
    sils = Split(ObjDTO.SilabasAuto, " | ")

    Dim i As Byte
    For i = LBound(sils) To UBound(sils)

        If Trim$(sils(i)) = "" Then
            ' Es un hueco
            If palabraActual.count > 0 Then
                resultado.Add palabraActual
                Set palabraActual = New Collection
            End If
            resultado.Add "HUECO"
        Else
            ' Es una sílaba real
            palabraActual.Add sils(i)
        End If

    Next i

    ' Última palabra
    If palabraActual.count > 0 Then resultado.Add palabraActual

    Set ObtenerPalabrasDesdeSilabasAuto = resultado

End Function


' ============================================================
'   JOIN COLLECTION
' ============================================================
Private Function JoinCollection(col As Collection) As String

    Dim arr() As String
    Dim i As Byte

    If col Is Nothing Then
        JoinCollection = ""
        Exit Function
    End If

    If col.count = 0 Then
        JoinCollection = ""
        Exit Function
    End If

    ReDim arr(1 To col.count)

    For i = 1 To col.count
        arr(i) = CStr(col(i))
    Next i

    JoinCollection = Join(arr, ",")

End Function



