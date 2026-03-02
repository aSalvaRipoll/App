Attribute VB_Name = "Temporal_2"

'Option Compare Database
'Option Explicit
'
'' ============================================================
''   MOTOR FONÉTICO — BALEAR (IB)
''   Construcción de IdsFonemas + cadena final
''   Arquitectura paralela al motor CA
'' ============================================================
'
'Public Sub ConstruirCadenaFonemas_IB()
'
'    Dim sils As Variant
'    Dim i As Long
'    Dim fonemaSils() As String
'    Dim cadena As String
'
'    ' 1) Obtener sílabas finales (ya acentuadas)
'    sils = Split(ObjDTO.SilabasFinal, " | ")
'
'    ReDim fonemaSils(LBound(sils) To UBound(sils))
'
'    ' 2) Convertir cada sílaba a fonemas (IDs)
'    For i = LBound(sils) To UBound(sils)
'        fonemaSils(i) = ConvertirSilaba_IB(Trim$(sils(i)))
'    Next i
'
'    ' 3) Unir sílabas fonéticas con separador
'    cadena = Join(fonemaSils, " . ")
'
'    ' 4) Procesos fonéticos globales (mismo orden que CA)
'    cadena = AplicarNeutralizaciones_IB(cadena)
'    cadena = AplicarAssimilaciones_IB(cadena)
'    cadena = AplicarReducciones_IB(cadena)
'    cadena = AplicarSchwa_IB(cadena)
'
'    ' 5) Guardar en DTO
'    ObjDTO.IdsFonemas = cadena
'    ObjDTO.FonemasFinal = cadena
'
'End Sub
'
'' ============================================================
''   CONVERTIR UNA SÍLABA IB A FONEMAS (IDs)
'' ============================================================
'Private Function ConvertirSilaba_IB(sil As String) As String
'
'    Dim out As String
'    Dim i As Long
'    Dim c As String, c2 As String
'    Dim grup As String
'
'    i = 1
'    Do While i <= Len(sil)
'
'        c = Mid$(sil, i, 1)
'
'        ' 1) Diptongos IB
'        If i < Len(sil) Then
'            c2 = Mid$(sil, i + 1, 1)
'            If EsDiptongo_IB(c, c2) Then
'                out = out & ConvertirDiptongo_IB(c, c2) & " "
'                i = i + 2
'                GoTo Siguiente
'            End If
'        End If
'
'        ' 2) Grupos consonánticos IB
'        If i < Len(sil) Then
'            grup = Mid$(sil, i, 2)
'            If EsGrupConsonantico_IB(grup) Then
'                out = out & ConvertirGrupo_IB(grup) & " "
'                i = i + 2
'                GoTo Siguiente
'            End If
'        End If
'
'        ' 3) Vocal sola / consonante sola
'        If EsVocal_IB(c) Then
'            out = out & ConvertirVocal_IB(c) & " "
'        Else
'            out = out & ConvertirConsonante_IB(c) & " "
'        End If
'
'Siguiente:
'        i = i + 1
'    Loop
'
'    ConvertirSilaba_IB = Trim$(out)
'
'End Function
'
'' ============================================================
''   CONVERTIR VOCAL IB ? ID FONEMA
'' ============================================================
'Private Function ConvertirVocal_IB(v As String) As String
'    Select Case v
'        Case "a", "à", "á"
'            ConvertirVocal_IB = "1"     ' /a/
'
'        Case "e", "é"
'            ConvertirVocal_IB = "2"     ' /e/
'
'        Case "è"
'            ConvertirVocal_IB = "3"     ' /?/
'
'        Case "i", "í", "ï"
'            ConvertirVocal_IB = "4"     ' /i/
'
'        Case "o", "ó"
'            ConvertirVocal_IB = "5"     ' /o/
'
'        Case "ò"
'            ConvertirVocal_IB = "6"     ' /?/
'
'        Case "u", "ú", "ü"
'            ConvertirVocal_IB = "7"     ' /u/
'
'        Case Else
'            ConvertirVocal_IB = "0"
'    End Select
'End Function
'
'' ============================================================
''   CONVERTIR CONSONANTE IB ? ID FONEMA
'' ============================================================
'Private Function ConvertirConsonante_IB(c As String) As String
'    Select Case c
'
'        Case "p": ConvertirConsonante_IB = "30"
'        Case "b": ConvertirConsonante_IB = "31"
'        Case "t": ConvertirConsonante_IB = "32"
'        Case "d": ConvertirConsonante_IB = "33"
'
'        Case "k", "c", "q"
'            ConvertirConsonante_IB = "34"   ' /k/
'
'        Case "g"
'            ConvertirConsonante_IB = "35"   ' /g/
'
'        Case "f": ConvertirConsonante_IB = "40"
'        Case "v": ConvertirConsonante_IB = "41"
'
'        Case "s": ConvertirConsonante_IB = "42"
'        Case "z": ConvertirConsonante_IB = "43"
'
'        Case "m": ConvertirConsonante_IB = "36"
'        Case "n": ConvertirConsonante_IB = "37"
'        Case "l": ConvertirConsonante_IB = "62"
'
'        Case "r"
'            ConvertirConsonante_IB = "59"   ' /?/
'
'        Case Else
'            ConvertirConsonante_IB = "0"
'    End Select
'End Function
'
'' ============================================================
''   GRUPOS CONSONÁNTICOS IB
'' ============================================================
'Private Function EsGrupConsonantico_IB(g As String) As Boolean
'    Select Case g
'        Case "ny", "ll", "rr", "ss", _
'             "tx", "tg", "tj", "ts", "tz", "ix"
'            EsGrupConsonantico_IB = True
'        Case Else
'            EsGrupConsonantico_IB = False
'    End Select
'End Function
'
'Private Function ConvertirGrupo_IB(g As String) As String
'    Select Case g
'
'        Case "ny": ConvertirGrupo_IB = "38"   ' /?/
'        Case "ll": ConvertirGrupo_IB = "63"   ' /?/
'        Case "rr": ConvertirGrupo_IB = "60"   ' /r/
'        Case "ss": ConvertirGrupo_IB = "42"   ' /s/
'
'        Case "tx": ConvertirGrupo_IB = "57"   ' /t?/
'        Case "tg", "tj": ConvertirGrupo_IB = "58"   ' /d?/
'
'        Case "ts", "tz": ConvertirGrupo_IB = "46"   ' /t?s/
'
'        Case "ix": ConvertirGrupo_IB = "44"   ' /?/
'
'        Case Else
'            ConvertirGrupo_IB = "0"
'    End Select
'End Function
'
'' ============================================================
''   CONVERTIR DIPTONGO IB ? IDs FONEMA
'' ============================================================
'Private Function ConvertirDiptongo_IB(c1 As String, c2 As String) As String
'    ConvertirDiptongo_IB = ConvertirVocal_IB(c1) & " " & ConvertirVocal_IB(c2)
'End Function
'
'' ============================================================
''   NEUTRALIZACIONES IB (DE MOMENTO VACÍO)
'' ============================================================
'Private Function AplicarNeutralizaciones_IB(cadena As String) As String
'    AplicarNeutralizaciones_IB = cadena
'End Function
'
'' ============================================================
''   ASSIMILACIONES IB
'' ============================================================
'Private Function AplicarAssimilaciones_IB(cadena As String) As String
'
'    ' S INTERVOCÁLICA ? Z (42 ? 43) según vocal siguiente
'    cadena = Replace(cadena, "42 1", "43 1")
'    cadena = Replace(cadena, "42 2", "43 2")
'    cadena = Replace(cadena, "42 3", "43 3")
'    cadena = Replace(cadena, "42 4", "43 4")
'    cadena = Replace(cadena, "42 5", "43 5")
'    cadena = Replace(cadena, "42 6", "43 6")
'    cadena = Replace(cadena, "42 7", "43 7")
'    cadena = Replace(cadena, "42 8", "43 8")
'
'    ' N + K/G ? ? (37 + 34/35 ? 39)
'    cadena = Replace(cadena, "37 34", "39")   ' nk
'    cadena = Replace(cadena, "37 35", "39")   ' ng
'
'    ' R INICIAL ? RR (59 ? 60)
'    If Left$(Trim$(cadena), 2) = "59" Then
'        cadena = "60" & Mid$(Trim$(cadena), 3)
'    End If
'
'    AplicarAssimilaciones_IB = cadena
'
'End Function
'
'' ============================================================
''   REDUCCIONES IB (LIMPIEZA FINAL)
'' ============================================================
'Private Function AplicarReducciones_IB(cadena As String) As String
'
'    Do While InStr(cadena, "  ") > 0
'        cadena = Replace(cadena, "  ", " ")
'    Loop
'
'    AplicarReducciones_IB = Trim$(cadena)
'
'End Function
'
'' ============================================================
''   SCHWA IB (? ? ID 8)
'' ============================================================
'Private Function AplicarSchwa_IB(cadena As String) As String
'
'    ' ARTICLE SALAT
'    cadena = Replace(cadena, "es ", "8 42 ")
'    cadena = Replace(cadena, "sa ", "42 8 ")
'    cadena = Replace(cadena, "ses ", "42 8 42 ")
'
'    ' APÒCOPES
'    cadena = Replace(cadena, "can' ", "34 8 37 ")
'    cadena = Replace(cadena, "ca' ", "34 8 ")
'
'    ' ALTRES PROCLÍTICS
'    cadena = Replace(cadena, "de ", "33 8 ")
'    cadena = Replace(cadena, "me ", "36 8 ")
'    cadena = Replace(cadena, "te ", "32 8 ")
'    cadena = Replace(cadena, "se ", "42 8 ")
'
'    AplicarSchwa_IB = cadena
'
'End Function
'
'
