Attribute VB_Name = "Módulo8"

'Option Compare Database
'Option Explicit
'
'Private modoDebug As Boolean   ' Activa / desactiva tracking
'
''----------------------------------------------------
'' FUNCIÓN PRINCIPAL
''----------------------------------------------------
'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim T As String
'    Dim nucIni() As Byte, nucFin() As Byte
'    Dim silIni() As Byte, silFin() As Byte
'    Dim nNuc As Byte, i As Byte
'    Dim silabas() As String
'
'    modoDebug = True
'
'    T = LCase$(Trim$(Texto))
'    If T = "" Then
'        SilabearPalabra = ""
'        Exit Function
'    End If
'
'    If modoDebug Then
'        addLog
'        addLog "============================================================="
'        addLog "        LOG DE DEPURACIÓN SILABEO ORTOGRÁFICO"
'        addLog "============================================================="
'        addLog "Entrada: '" & Texto & "'"
'        addLog "Normalizado: '" & T & "'"
'        addLog "Longitud: " & Len(T) & " letras."
'    End If
'
'    ' 1. Localizar núcleos ortográficos (con diptongos/triptongos)
'    LocalizarNucleosOrtog T, nucIni, nucFin, nNuc
'
'    If modoDebug Then
'        Dim tmp As String
'        tmp = ""
'        For i = 1 To nNuc
'            If tmp <> "" Then tmp = tmp & ","
'            tmp = tmp & CStr(nucIni(i)) & "-" & CStr(nucFin(i))
'        Next i
'        addLog "Núcleos (ini-fin): " & tmp
'    End If
'
'    ' 2. Calcular límites de sílabas según reglas ortográficas
'    ReDim silIni(1 To nNuc)
'    ReDim silFin(1 To nNuc)
'
'    CalcularSilabas T, nucIni, nucFin, nNuc, silIni, silFin
'
'    ' 3. Construir sílabas
'    ReDim silabas(1 To nNuc)
'    For i = 1 To nNuc
'        silabas(i) = Mid$(T, silIni(i), silFin(i) - silIni(i) + 1)
'        If modoDebug Then
'            addLog "Sílaba " & i & ": [" & silIni(i) & "-" & silFin(i) & "] = " & silabas(i)
'        End If
'    Next i
'
'    SilabearPalabra = Join(silabas, " | ")
'
'    If modoDebug Then
'        addLog
'        addLog "Resultado: " & SilabearPalabra
'        addLog "============================================================="
'        addLog "                   FIN LOG DE DEPURACIÓN"
'        addLog "============================================================="
'        PrintLog
'    End If
'End Function
'
''----------------------------------------------------
'' LOCALIZAR NÚCLEOS (ORTOGRÁFICO: Diptongos/Triptongos)
''----------------------------------------------------
'Private Sub LocalizarNucleosOrtog(ByVal T As String, _
'                                  ByRef nucIni() As Byte, _
'                                  ByRef nucFin() As Byte, _
'                                  ByRef nNuc As Byte)
'    Dim i As Byte, L As Byte
'    Dim C1 As String, C2 As String, C3 As String
'
'    L = Len(T)
'    ReDim nucIni(1 To L)
'    ReDim nucFin(1 To L)
'    nNuc = 0
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento LocalizarNucleosOrtog"
'    End If
'
'    i = 1
'    Do While i <= L
'        C1 = Mid$(T, i, 1)
'        If EsVocal(C1) Then
'            ' Intentar triptongo
'            If i + 2 <= L Then
'                C2 = Mid$(T, i + 1, 1)
'                C3 = Mid$(T, i + 2, 1)
'                If EsTriptongo(C1, C2, C3) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 2
'                    If modoDebug Then addLog "Triptongo en " & i & "-" & i + 2 & " (" & C1 & C2 & C3 & ")"
'                    i = i + 3
'                    GoTo Siguiente
'                End If
'            End If
'            ' Intentar diptongo
'            If i + 1 <= L Then
'                C2 = Mid$(T, i + 1, 1)
'                If EsVocal(C2) And EsDiptongo(C1, C2) Then
'                    nNuc = nNuc + 1
'                    nucIni(nNuc) = i
'                    nucFin(nNuc) = i + 1
'                    If modoDebug Then addLog "Diptongo en " & i & "-" & i + 1 & " (" & C1 & C2 & ")"
'                    i = i + 2
'                    GoTo Siguiente
'                End If
'            End If
'            ' Vocal sola
'            nNuc = nNuc + 1
'            nucIni(nNuc) = i
'            nucFin(nNuc) = i
'            If modoDebug Then addLog "Vocal sola en " & i & " (" & C1 & ")"
'            i = i + 1
'        Else
'            i = i + 1
'        End If
'Siguiente:
'    Loop
'
'    If modoDebug Then
'        addLog "Total núcleos: " & nNuc
'        addLog " Fin LocalizarNucleosOrtog"
'        addLog "---------------------------------------"
'    End If
'End Sub
'
''----------------------------------------------------
'' CÁLCULO DE SÍLABAS (REGLAS ORTOGRÁFICAS)
''----------------------------------------------------
'Private Sub CalcularSilabas(ByVal T As String, _
'                            ByRef nucIni() As Byte, _
'                            ByRef nucFin() As Byte, _
'                            ByVal nNuc As Byte, _
'                            ByRef silIni() As Byte, _
'                            ByRef silFin() As Byte)
'    Dim i As Byte, L As Byte
'    Dim a As Byte, b As Byte
'    Dim k As Byte
'    Dim C1 As String, C2 As String, C3 As String, C4 As String
'    Dim grupo As String
'
'    L = Len(T)
'
'    ' Primera sílaba empieza al principio
'    silIni(1) = 1
'
'    For i = 1 To nNuc - 1
'        a = nucFin(i)
'        b = nucIni(i + 1)
'
'        k = IIf(b > a + 1, b - a - 1, 0)   ' nº consonantes entre núcleos
'
'        If modoDebug Then
'            addLog
'            addLog "---- Frontera entre núcleo " & i & " (" & a & ") y " & (i + 1) & " (" & b & ")"
'            addLog "Consonantes entre medias: " & k
'        End If
'
'        Select Case k
'            Case 0
'                ' V (hiato) V ? V | V
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'            Case 1
'                ' V C V ? V | CV
'                silFin(i) = a
'                silIni(i + 1) = a + 1
'            Case 2
'                ' V C1 C2 V
'                C1 = Mid$(T, a + 1, 1)
'                C2 = Mid$(T, a + 2, 1)
'                grupo = C1 & C2
'                If modoDebug Then addLog "Grupo CC: " & grupo
'                If EsGrupoAtaque(grupo) Then
'                    ' V | CCV
'                    silFin(i) = a
'                    silIni(i + 1) = a + 1
'                Else
'                    ' VC | CV
'                    silFin(i) = a + 1
'                    silIni(i + 1) = a + 2
'                End If
'            Case 3
'                ' V C1 C2 C3 V ? V C | CCV
'                silFin(i) = a + 1
'                silIni(i + 1) = a + 2
'            Case Else
'                ' V CCCC V ? V CC | CCV
'                silFin(i) = a + 2
'                silIni(i + 1) = a + 3
'        End Select
'
'        If modoDebug Then
'            addLog "silFin(" & i & ") = " & silFin(i)
'            addLog "silIni(" & (i + 1) & ") = " & silIni(i + 1)
'        End If
'    Next i
'
'    ' Última sílaba termina al final de la palabra
'    silFin(nNuc) = L
'End Sub
'
''----------------------------------------------------
'' VOCAL / DÍPTONGO / TRIPTONGO
''----------------------------------------------------
'Private Function EsVocal(ByVal c As String) As Boolean
'    EsVocal = InStr("aeiouáéíóúü", c) > 0
'End Function
'
'Private Function EsVocalFuerte(ByVal c As String) As Boolean
'    EsVocalFuerte = InStr("aáeéoó", c) > 0
'End Function
'
'Private Function EsVocalDebil(ByVal c As String) As Boolean
'    EsVocalDebil = InStr("iíuúü", c) > 0
'End Function
'
'Private Function EsVocalDebilTonica(ByVal c As String) As Boolean
'    EsVocalDebilTonica = InStr("íú", c) > 0
'End Function
'
'Private Function EsDiptongo(ByVal v1 As String, ByVal v2 As String) As Boolean
'    ' Regla simplificada RAE:
'    ' - dos débiles ? diptongo (iu, ui, üi, güe, etc.)
'    ' - fuerte + débil átona ? diptongo
'    ' - débil átona + fuerte ? diptongo
'    ' - dos fuertes ? hiato
'    ' - débil tónica + fuerte ? hiato
'    ' - fuerte + débil tónica ? hiato
'    If Not EsVocal(v1) Or Not EsVocal(v2) Then
'        EsDiptongo = False
'        Exit Function
'    End If
'
'    If EsVocalFuerte(v1) And EsVocalFuerte(v2) Then
'        EsDiptongo = False
'        Exit Function
'    End If
'
'    If EsVocalDebilTonica(v1) Or EsVocalDebilTonica(v2) Then
'        EsDiptongo = False
'        Exit Function
'    End If
'
'    EsDiptongo = True
'End Function
'
'Private Function EsTriptongo(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As Boolean
'    ' débil átona + fuerte + débil átona
'    If EsVocalDebil(v1) And Not EsVocalDebilTonica(v1) _
'       And EsVocalFuerte(v2) _
'       And EsVocalDebil(v3) And Not EsVocalDebilTonica(v3) Then
'        EsTriptongo = True
'    Else
'        EsTriptongo = False
'    End If
'End Function
'
''----------------------------------------------------
'' GRUPOS CONSONÁNTICOS DE ATAQUE (ORTOGRÁFICO)
''----------------------------------------------------
'Private Function EsGrupoAtaque(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "gr", "fr", _
'               "pl", "bl", "cl", "gl", "fl")
'    EsGrupoAtaque = (UBound(Filter(AC, g)) >= 0)
'End Function
'
