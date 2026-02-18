Attribute VB_Name = "modMotorFonetico_V2_1_Aux"
Option Compare Database
Option Explicit



Public Function MF_Silabear(ByVal Texto As String, ByVal Abreviado As String) As Collection

    Dim silabas As Collection

    ' ============================================================
    ' 1. REGLAS UNIVERSALES (comunes a todos los idiomas)
    ' ============================================================
    ' - Definición universal de vocales
    ' - Diptongos universales
    ' - Hiatos universales
    ' - División CV universal
    ' - Manejo de consonantes
    ' - Estructura base de sílaba
    '
    ' Estas reglas se aplican SIEMPRE.
    ' ============================================================

    ' Aquí llamamos al silabeador base universal
    Set silabas = MF_SilabearUniversalBase(Texto)

    ' ============================================================
    ' 2. MICROAJUSTES POR IDIOMA
    ' ============================================================
    Select Case LCase$(Abreviado)

        ' --------------------------------------------------------
        ' CATALÁN / VALENCIANO / MALLORQUÍN
        ' --------------------------------------------------------
        Case "ca", "ca-va", "ca-ib"
            Call MF_SilabearAjustesCatalan(Texto, silabas)

        ' --------------------------------------------------------
        ' CASTELLANO / GALLEGO
        ' --------------------------------------------------------
        Case "es", "gl"
            Call MF_SilabearAjustesCastellanoGallego(Texto, silabas)

        ' --------------------------------------------------------
        ' PORTUGUÉS EUROPEO / BRASILEÑO
        ' --------------------------------------------------------
        Case "pt", "br"
            Call MF_SilabearAjustesPortugues(Texto, silabas)

        ' --------------------------------------------------------
        ' FRANCÉS
        ' --------------------------------------------------------
        Case "fr"
            Call MF_SilabearAjustesFrances(Texto, silabas)

        ' --------------------------------------------------------
        ' INGLÉS
        ' --------------------------------------------------------
        Case "en"
            Call MF_SilabearAjustesIngles(Texto, silabas)

        ' --------------------------------------------------------
        ' EUSKARA
        ' --------------------------------------------------------
        Case "eu"
            Call MF_SilabearAjustesEuskara(Texto, silabas)

    End Select

    Set MF_Silabear = silabas

End Function

Public Function MF_SilabearUniversalBase(ByVal Texto As String) As Collection

    Dim col As New Collection
    Dim i As Long, ini As Long

    ' Vocales universales (romance + francés + portugués + catalán + euskara)
    Dim vocales As String
    vocales = "AEIOU" & _
              "ÁÉÍÓÚ" & _
              "ÀÈÌÒÙ" & _
              "ÂÊÔ" & _
              "ÃÕ" & _
              "Ü"

    ' Vocales débiles universales
    Dim vocalesDebiles As String
    vocalesDebiles = "IÍÌUÚÙÜ"

    ' Vocales fuertes universales
    Dim vocalesFuertes As String
    vocalesFuertes = "AÁÀÂÃEÉÈÊOÓÒÔÕ"

    ini = 1

    For i = 2 To Len(Texto)

        Dim c1 As String, c2 As String
        c1 = Mid$(Texto, i - 1, 1)
        c2 = Mid$(Texto, i, 1)

        Dim esV1 As Boolean, esV2 As Boolean
        esV1 = (InStr(vocales, c1) > 0)
        esV2 = (InStr(vocales, c2) > 0)

        ' --------------------------------------------------------
        ' Caso 1: Vocal + Vocal ? decidir diptongo o hiato
        ' --------------------------------------------------------
        If esV1 And esV2 Then

            Dim diptongo As Boolean
            diptongo = False

            ' Diptongo universal:
            ' fuerte+débil, débil+fuerte, débil+débil
            If (InStr(vocalesFuertes, c1) > 0 And InStr(vocalesDebiles, c2) > 0) _
            Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesFuertes, c2) > 0) _
            Or (InStr(vocalesDebiles, c1) > 0 And InStr(vocalesDebiles, c2) > 0) Then
                diptongo = True
            End If

            ' Si NO hay diptongo ? cerrar sílaba
            If Not diptongo Then
                col.Add Array(ini, i - 1)
                ini = i
            End If

        End If

        ' --------------------------------------------------------
        ' Caso 2: Consonante ? Vocal
        ' (posible inicio de sílaba, pero no cerramos aún)
        ' --------------------------------------------------------
        ' No hacemos nada aquí

        ' --------------------------------------------------------
        ' Caso 3: Vocal ? Consonante
        ' (posible cierre, pero esperamos a la siguiente vocal)
        ' --------------------------------------------------------
        ' No hacemos nada aquí

    Next i

    ' Última sílaba
    col.Add Array(ini, Len(Texto))

    Set MF_SilabearUniversalBase = col

End Function

Public Sub MF_SilabearAjustesCastellanoGallego( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long
    Dim vocalesTilde As String
    vocalesTilde = "ÁÉÍÓÚ"

    ' ============================================================
    ' 1. HIATOS CON TILDE (rompen diptongo)
    ' ============================================================
    For i = 2 To Len(Texto)
        If InStr(vocalesTilde, Mid$(Texto, i, 1)) > 0 Then

            ' Si la vocal anterior es vocal ? romper sílaba
            If InStr("AEIOUÁÉÍÓÚ", Mid$(Texto, i - 1, 1)) > 0 Then
                Call MF_ForzarDivisionSilabica(silabas, i)
            End If

        End If
    Next i

    ' ============================================================
    ' 2. GRUPOS CONSONÁNTICOS INSEPARABLES
    ' ============================================================
    Dim grupos As Variant
    grupos = Array("BR", "BL", "CR", "CL", "DR", "FR", "FL", _
                   "GR", "GL", "PR", "PL", "TR", "TL")

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, grupos) Then
            ' Si el universal ha dividido entre estas dos consonantes ? corregir
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub

Private Sub MF_ForzarDivisionSilabica(ByRef silabas As Collection, ByVal pos As Long)

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

Private Sub MF_UnirConsonantesEnAtaque(ByRef silabas As Collection, ByVal pos As Long)

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        ' Si la división está justo antes de pos ? unir
        If fin = pos - 1 Then
            ' Unir esta sílaba con la siguiente
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

Private Function EsMiembro(ByVal s As String, ByVal arr As Variant) As Boolean
    Dim x As Variant
    For Each x In arr
        If s = x Then
            EsMiembro = True
            Exit Function
        End If
    Next x
    EsMiembro = False
End Function
'--------------------------------------------------------------------------------------

Public Sub MF_SilabearAjustesCatalanCentral( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL CATALÁN CENTRAL
    ' ============================================================
    ' Nota: estos grupos deben permanecer juntos si van seguidos de vocal.
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR", "TL", "DL", _
        "SC", "SP", "ST", "SM", "SN" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. Diptongos catalanes (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO", _
        "AI", "EI", "OI", "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim seq As String
        seq = Mid$(Texto, i - 1, 2)

        If EsMiembro(seq, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub

Private Sub MF_UnirConsonantesEnAtaqueCA(ByRef silabas As Collection, ByVal pos As Long)

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

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

Private Sub MF_UnirVocalesEnDiptongo(ByRef silabas As Collection, ByVal pos As Long)

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

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

Private Function EsMiembroCA(ByVal s As String, ByVal arr As Variant) As Boolean
    Dim x As Variant
    For Each x In arr
        If s = x Then
            EsMiembro = True
            Exit Function
        End If
    Next x
    EsMiembro = False
End Function

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesValenciano( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL VALENCIANO
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR", "TL", "DL", _
        "SC", "SP", "ST", "SM", "SN" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. TS y TZ ? no dividir si van seguidas de vocal
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "TS" Or seq = "TZ" Then
            If EsVocal(Mid$(Texto, i + 2, 1)) Then
                Call MF_UnirConsonantesEnAtaque(silabas, i)
            End If
        End If
    Next i

    ' ============================================================
    ' 3. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS VALENCIANOS (ea, eo, oa)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("EA", "EO", "OA")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. Diptongos valencianos (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub

Private Function EsVocalMF(ByVal c As String) As Boolean
    EsVocal = InStr("AEIOUÁÉÍÓÚÀÈÌÒÙÂÊÔÃÕÜ", c) > 0
End Function

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesMallorquin( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL MALLORQUÍN
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR", "TL", "DL", _
        "SC", "SP", "ST", "SM", "SN" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. LL y RR nunca se separan
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "LL" Or Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. HIATOS MALLORQUINES (ea, eo, oa)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("EA", "EO", "OA")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. Diptongos mallorquines (muy estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO", _
        "AI", "EI", "OI", "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

End Sub


'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesPortuguesBR( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PERMITIDOS EN PT-BR
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. DIPTONGOS NASAIS (muy estables)
    ' ============================================================
    Dim diptNasal As Variant
    diptNasal = Array("ÃO", "ÃE", "ÕE", "ÕI", "ÃI", "ÕU")

    For i = 2 To Len(Texto)
        Dim dn As String
        dn = Mid$(Texto, i - 1, 2)

        If EsMiembro(dn, diptNasal) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. DIPTONGOS ORALES (muy estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS OBLIGATORIOS (aí, aí, oá, oé, oê, eí, eú…)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("AÍ", "AÌ", "AÚ", "AÓ", _
                   "EÍ", "EÚ", "OÁ", "OÉ", "OÊ", _
                   "IÁ", "IÉ", "IÓ", "IÚ")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. RR y SS intervocálicas ? no dividir
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "RR" Or seq = "SS" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesPortuguesEU( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PERMITIDOS EN PT-PT
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "CR", "CL", "DR", "FR", "FL", _
        "GR", "GL", "PR", "PL", "TR" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. DIPTONGOS NASAIS (PT-PT)
    ' ============================================================
    Dim diptNasal As Variant
    diptNasal = Array("ÃO", "ÃE", "ÕE", "ÕI")

    For i = 2 To Len(Texto)
        Dim dn As String
        dn = Mid$(Texto, i - 1, 2)

        If EsMiembro(dn, diptNasal) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. VOCALES NASAIS FINALES (am, em, im, om, um)
    '    ? deben ser UNA sola sílaba
    ' ============================================================
    Dim nasalesFinales As Variant
    nasalesFinales = Array("AM", "EM", "IM", "OM", "UM")

    For i = 2 To Len(Texto)
        Dim nf As String
        nf = Mid$(Texto, i - 1, 2)

        If EsMiembro(nf, nasalesFinales) And i = Len(Texto) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS PT-PT (ea, eo, oa, oe, ui, iu)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("EA", "EO", "OA", "OE", "UI", "IU")

    For i = 2 To Len(Texto)
        Dim hv As String
        hv = Mid$(Texto, i - 1, 2)

        If EsMiembro(hv, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 5. DIPTONGOS ORALES PT-PT (estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 6. RR y SS intervocálicas ? no dividir
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If seq = "RR" Or seq = "SS" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesFrances( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ELIMINAR LA E MUDA FINAL
    ' ============================================================
    If Right$(Texto, 1) = "E" Then
        Call MF_EliminarVocalFinal(silabas, Len(Texto))
    End If

    ' ============================================================
    ' 2. AGRUPAR NASALES (an, am, en, em, in, im, ain, ein, un, um, on, om)
    ' ============================================================
    Dim nasales As Variant
    nasales = Array("AN", "AM", "EN", "EM", "IN", "IM", "AIN", "EIN", "UN", "UM", "ON", "OM")

    For i = 2 To Len(Texto)
        Dim seq2 As String, seq3 As String
        seq2 = Mid$(Texto, i - 1, 2)
        seq3 = Mid$(Texto, i - 2, 3)

        If EsMiembro(seq3, nasales) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i - 1)
        ElseIf EsMiembro(seq2, nasales) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. ATAQUES CONSONÁNTICOS PERMITIDOS EN FRANCÉS
    ' ============================================================
    Dim ataques As Variant
    ataques = Array("TR", "DR", "PR", "BR", "CR", "GR", "FR", "FL", "CL", "GL", "PL")

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. HIATOS POR DIÉRESIS (ï, ë, ü)
    ' ============================================================
    Dim dieresis As String
    dieresis = "ÏËÜ"

    For i = 2 To Len(Texto)
        If InStr(dieresis, Mid$(Texto, i, 1)) > 0 Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub

Private Sub MF_EliminarVocalFinal(ByRef silabas As Collection, ByVal pos As Long)

    Dim i As Long

    For i = 1 To silabas.Count
        Dim ini As Long, fin As Long
        ini = silabas(i)(0)
        fin = silabas(i)(1)

        If fin = pos Then
            ' Reducir la sílaba eliminando la E final
            silabas.Remove i
            silabas.Add Array(ini, pos - 1), , i
            Exit Sub
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesIngles( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PROPIOS DEL INGLÉS
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "PL", "PR", "BL", "BR", "CL", "CR", "GL", "GR", _
        "FL", "FR", "TR", "DR", "SK", "SL", "SM", "SN", _
        "SP", "ST", "SW", "SH", "TH", "WH", _
        "STR", "SPR", "SPL", "SCR", "SHR", "THR" _
    )

    ' Primero ataques de 3 letras
    For i = 3 To Len(Texto) - 1
        Dim tri As String
        tri = Mid$(Texto, i - 1, 3)

        If EsMiembro(tri, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i - 1)
        End If
    Next i

    ' Luego ataques de 2 letras
    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. DIPTONGOS INGLESES (refuerzo)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "AY", "EI", "EY", "OI", "OY", _
        "AU", "AW", "OU", "OW", _
        "EA", "EE", "IE", "OA" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. HIATOS OBLIGATORIOS (cooperate, naive, reenter…)
    ' ============================================================
    Dim hiatos As Variant
    hiatos = Array("COO", "OO", "RE", "NAI", "REI")

    For i = 3 To Len(Texto)
        Dim tri2 As String
        tri2 = Mid$(Texto, i - 2, 3)

        If EsMiembro(tri2, hiatos) Then
            Call MF_ForzarDivisionSilabica(silabas, i - 1)
        End If
    Next i

    ' ============================================================
    ' 4. DIÉRESIS (ï, ë) ? rompen diptongo
    ' ============================================================
    Dim dieresis As String
    dieresis = "ÏË"

    For i = 2 To Len(Texto)
        If InStr(dieresis, Mid$(Texto, i, 1)) > 0 Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------
Public Sub MF_SilabearAjustesEuskara( _
        ByVal Texto As String, _
        ByRef silabas As Collection _
    )

    Dim i As Long

    ' ============================================================
    ' 1. ATAQUES CONSONÁNTICOS PERMITIDOS EN EUSKARA
    ' ============================================================
    Dim ataques As Variant
    ataques = Array( _
        "BR", "BL", "TR", "DR", "KR", "KL", _
        "GR", "GL", "PR", "PL" _
    )

    For i = 2 To Len(Texto) - 1
        Dim par As String
        par = Mid$(Texto, i, 2)

        If EsMiembro(par, ataques) Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 2. RR nunca se separa
    ' ============================================================
    For i = 2 To Len(Texto) - 1
        If Mid$(Texto, i, 2) = "RR" Then
            Call MF_UnirConsonantesEnAtaque(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 3. DIPTONGOS EUSKÉRICOS (muy estables)
    ' ============================================================
    Dim dipt As Variant
    dipt = Array( _
        "AI", "EI", "OI", "UI", _
        "AU", "EU", "OU", _
        "IA", "IE", "IO", "IU", _
        "UA", "UE", "UI", "UO" _
    )

    For i = 2 To Len(Texto)
        Dim dv As String
        dv = Mid$(Texto, i - 1, 2)

        If EsMiembro(dv, dipt) Then
            Call MF_UnirVocalesEnDiptongo(silabas, i)
        End If
    Next i

    ' ============================================================
    ' 4. EVITAR ATAQUES NO PERMITIDOS (SC, SP, ST, SM, SN, TL, DL…)
    '    ? si el universal los ha unido, los separamos
    ' ============================================================
    Dim noAtaques As Variant
    noAtaques = Array("SC", "SP", "ST", "SM", "SN", "TL", "DL", "TS", "TX", "TZ")

    For i = 2 To Len(Texto) - 1
        Dim seq As String
        seq = Mid$(Texto, i, 2)

        If EsMiembro(seq, noAtaques) Then
            Call MF_ForzarDivisionSilabica(silabas, i)
        End If
    Next i

End Sub

'-----------------------------------------------------------------------------------

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



Public Function MF_NormalizarVocalesPorIdioma( _
    ByVal Texto As String, _
    ByVal idioma As String _
    ) As String

    Select Case UCase$(idioma)

        Case "ES"
            Texto = MF_NormalizarVocales_ES(Texto)

        Case "CA" ', "CA-IB", "CA-VA"
            Texto = MF_NormalizarVocales_CA(Texto)
            
        Case "CA-IB"
            Texto = MF_NormalizarVocales_CA_IB(Texto)

        Case "CA-VA"
            Texto = MF_NormalizarVocales_CA_VA(Texto)
            
        Case "GL"
            Texto = MF_NormalizarVocales_GL(Texto)

        Case "EU"
            Texto = MF_NormalizarVocales_EU(Texto)

        Case "PT-EU"
            Texto = MF_NormalizarVocales_PT_EU(Texto)
        
        Case "PT-BR"
            Texto = MF_NormalizarVocales_PT_BR(Texto)

        Case "FR"
            Texto = MF_NormalizarVocales_FR(Texto)

        Case "EN-GB"
            Texto = MF_NormalizarVocales_EN_GB(Texto)

'        Case "EN-US"
'            texto = MF_NormalizarVocales_EN_US(texto)

'        Case "EN-US-AF"
'            texto = MF_NormalizarVocales_EN_US_AF(texto)

        Case Else
            Texto = MF_NormalizarVocales_General(Texto)

    End Select

    MF_NormalizarVocalesPorIdioma = Texto

End Function

Private Function MF_NormalizarVocales_ES(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Ä", "A")
    Texto = Replace(Texto, "Â", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ë", "E")
    Texto = Replace(Texto, "Ê", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ï", "I")
    Texto = Replace(Texto, "Î", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ö", "O")
    Texto = Replace(Texto, "Ô", "O")

    ' U (sin tocar Ü)
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")
    Texto = Replace(Texto, "Û", "U")

    MF_NormalizarVocales_ES = Texto

End Function

Private Function MF_NormalizarVocales_CA(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Á", "A")

    ' E
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "É", "E")

    ' I  (NO tocar Ï)
    Texto = Replace(Texto, "Í", "I")

    ' O
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ó", "O")

    ' U  (NO tocar Ü)
    Texto = Replace(Texto, "Ú", "U")

    MF_NormalizarVocales_CA = Texto

End Function

Private Function MF_NormalizarVocales_CA_IB(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_IB = MF_NormalizarVocales_CA(Texto)
End Function

Private Function MF_NormalizarVocales_CA_VA(ByVal Texto As String) As String
    MF_NormalizarVocales_CA_VA = MF_NormalizarVocales_CA(Texto)
End Function

Private Function MF_NormalizarVocales_GL(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")

    ' U
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_GL = Texto

End Function


Private Function MF_NormalizarVocales_EU(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")

    ' E
    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")

    ' I
    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")

    ' O
    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")

    ' U
    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_EU = Texto

End Function

Private Function MF_NormalizarVocales_PT_EU(ByVal Texto As String) As String

    ' Nasales
    Texto = Replace(Texto, "Ã", "A~")
    Texto = Replace(Texto, "Õ", "O~")

    ' Cerradas (circunflejo)
    Texto = Replace(Texto, "Â", "Â")
    Texto = Replace(Texto, "Ê", "Ê")
    Texto = Replace(Texto, "Î", "I") ' no existe en PT, pero por robustez
    Texto = Replace(Texto, "Ô", "Ô")
    Texto = Replace(Texto, "Û", "U") ' no existe en PT, robustez

    ' Abiertas (agudas)
    Texto = Replace(Texto, "Á", "A´")
    Texto = Replace(Texto, "É", "E´")
    Texto = Replace(Texto, "Í", "I´")
    Texto = Replace(Texto, "Ó", "O´")
    Texto = Replace(Texto, "Ú", "U´")

    ' Graves (no existen en PT, pero pueden aparecer en nombres importados)
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_PT_EU = Texto

End Function

Private Function MF_NormalizarVocales_PT_BR(ByVal Texto As String) As String

    ' Nasales (idénticas a PT-EU)
    Texto = Replace(Texto, "Ã", "A~")
    Texto = Replace(Texto, "Õ", "O~")

    ' Cerradas ? se suavizan (PT-BR no las mantiene tan tensas)
    Texto = Replace(Texto, "Â", "A")
    Texto = Replace(Texto, "Ê", "E")
    Texto = Replace(Texto, "Ô", "O")

    ' Abiertas (agudas)
    Texto = Replace(Texto, "Á", "A´")
    Texto = Replace(Texto, "É", "E´")
    Texto = Replace(Texto, "Í", "I´")
    Texto = Replace(Texto, "Ó", "O´")
    Texto = Replace(Texto, "Ú", "U´")

    ' Graves (robustez)
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ù", "U")

    MF_NormalizarVocales_PT_BR = Texto

End Function

Private Function MF_NormalizarVocales_FR(ByVal Texto As String) As String

    ' A
    Texto = Replace(Texto, "À", "A")   ' abierta
    Texto = Replace(Texto, "Á", "A")   ' rara, pero robustez
    Texto = Replace(Texto, "Â", "Â")   ' cerrada
    Texto = Replace(Texto, "Ä", "A¨")  ' hiato

    ' E
    Texto = Replace(Texto, "È", "E")   ' abierta
    Texto = Replace(Texto, "É", "E´")  ' cerrada
    Texto = Replace(Texto, "Ê", "Ê")   ' cerrada tensa
    Texto = Replace(Texto, "Ë", "E¨")  ' hiato

    ' I
    Texto = Replace(Texto, "Ì", "I")   ' robustez
    Texto = Replace(Texto, "Í", "I")   ' robustez
    Texto = Replace(Texto, "Î", "Î")   ' cerrada
    Texto = Replace(Texto, "Ï", "I¨")  ' hiato

    ' O
    Texto = Replace(Texto, "Ò", "O")   ' robustez
    Texto = Replace(Texto, "Ó", "O")   ' robustez
    Texto = Replace(Texto, "Ô", "Ô")   ' cerrada
    Texto = Replace(Texto, "Ö", "O¨")  ' hiato

    ' U
    Texto = Replace(Texto, "Ù", "U")   ' abierta
    Texto = Replace(Texto, "Ú", "U")   ' robustez
    Texto = Replace(Texto, "Û", "Û")   ' cerrada
    Texto = Replace(Texto, "Ü", "U¨")  ' hiato

    MF_NormalizarVocales_FR = Texto

End Function

Private Function MF_NormalizarVocales_EN_GB(ByVal Texto As String) As String

    ' Solo por robustez ante nombres importados
    Texto = Replace(Texto, "Á", "A")
    Texto = Replace(Texto, "À", "A")
    Texto = Replace(Texto, "Ä", "A")
    Texto = Replace(Texto, "Â", "A")

    Texto = Replace(Texto, "É", "E")
    Texto = Replace(Texto, "È", "E")
    Texto = Replace(Texto, "Ë", "E")
    Texto = Replace(Texto, "Ê", "E")

    Texto = Replace(Texto, "Í", "I")
    Texto = Replace(Texto, "Ì", "I")
    Texto = Replace(Texto, "Ï", "I")
    Texto = Replace(Texto, "Î", "I")

    Texto = Replace(Texto, "Ó", "O")
    Texto = Replace(Texto, "Ò", "O")
    Texto = Replace(Texto, "Ö", "O")
    Texto = Replace(Texto, "Ô", "O")

    Texto = Replace(Texto, "Ú", "U")
    Texto = Replace(Texto, "Ù", "U")
    Texto = Replace(Texto, "Ü", "U")
    Texto = Replace(Texto, "Û", "U")

    MF_NormalizarVocales_EN_GB = Texto

End Function

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

Private Function MF_NormalizarVocales_General(ByVal Texto As String) As String
    
    Dim i As Integer
    Dim c As String
    Dim Res As String

    c = ""
    Res = ""
    
    For i = 1 To Len(Texto)
        c = Mid(Texto, i, 1)
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
        
        Res = Res & c
    Next i
    
    MF_NormalizarVocales_General = Res

End Function

Public Function EsVocal(ByVal c As String) As Boolean

'Versión blindada UNICODE

    Dim code As Long
    code = AscW(c)

    ' Vocales básicas A E I O U
    If code = &H41 Or code = &H45 Or code = &H49 Or code = &H4F Or code = &H55 Then
        EsVocal = True: Exit Function
    End If

    ' Vocales minúsculas (por si acaso)
    If code = &H61 Or code = &H65 Or code = &H69 Or code = &H6F Or code = &H75 Then
        EsVocal = True: Exit Function
    End If

    ' Vocales acentuadas (agudas, graves, circunflejos, diéresis)
    ' Rango general: U+00C0 – U+00FF (letras latinas extendidas)
    If code >= &HC0 And code <= &HFF Then
        Select Case code
            ' Á É Í Ó Ú
            Case &HC1, &HC9, &HCD, &HD3, &HDA
                EsVocal = True: Exit Function

            ' À È Ì Ò Ù
            Case &HC0, &HC8, &HCC, &HD2, &HD9
                EsVocal = True: Exit Function

            ' Â Ê Î Ô Û
            Case &HC2, &HCA, &HCE, &HD4, &HDB
                EsVocal = True: Exit Function

            ' Ä Ë Ï Ö Ü
            Case &HC4, &HCB, &HCF, &HD6, &HDC
                EsVocal = True: Exit Function

            ' Nasales portuguesas: Ã Õ
            Case &HC3, &HD5
                EsVocal = True: Exit Function
        End Select
    End If

    ' Si no coincide con nada
    EsVocal = False
    
End Function

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

Public Function EsConsonante(ByVal c As String) As Boolean
    If c = "" Then
        EsConsonante = False
    Else
        EsConsonante = Not EsVocal(c)
    End If
End Function

Public Function ProcesarY(ByVal ant As String, ByVal sig As String) As String
    If (ant = " ") And (sig = " ") Then
        ProcesarY = "I": Exit Function
    End If

    If ant = "" And EsVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If ant = "" And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If sig = "" Then
        ProcesarY = "I": Exit Function
    End If

    If EsVocal(ant) And EsVocal(sig) Then
        ProcesarY = "Y": Exit Function
    End If

    If EsConsonante(ant) And EsConsonante(sig) Then
        ProcesarY = "I": Exit Function
    End If

    If EsVocal(sig) Then
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
        f.EsVocal = rs!EsVocal
        'f.Valor = rs!Valor
        ' ... cualquier otra propiedad

        colFonemas.Add f, CStr(f.idFonema)
        rs.MoveNext
    Loop

    rs.Close
End Sub

'Public Function BuscarExcepcion(ByVal palabra As String, ByVal idioma As String) As String
'
'    Dim rs As DAO.Recordset
'    Set rs = CurrentDb.OpenRecordset( _
'        "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
'        "WHERE Idioma = '" & idioma & "' AND Palabra = '" & UCase(palabra) & "' AND Activo = True")
'
'    If Not rs.EOF Then
'        BuscarExcepcion = rs!FonemaCompleto
'    Else
'        BuscarExcepcion = ""
'    End If
'
'    rs.Close
'    Set rs = Nothing
'
'End Function

Public Function BuscarExcepcion(ByVal graf As String, ByVal idioma As String) As Byte

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT idFonema FROM tbmDicExcepciones " & _
          "WHERE Tipo = 'GRAFEMA' " & _
          "AND Idioma = '" & idioma & "' " & _
          "AND Grafema = '" & graf & "' " & _
          "AND Activo = True"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        BuscarExcepcion = rs!idFonema
    Else
        BuscarExcepcion = 0
    End If

    rs.Close
    Set rs = Nothing

End Function

Public Function BuscarExcepcionPalabra(ByVal palabra As String, ByVal idioma As String) As String

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT FonemaCompleto FROM tbmDicExcepciones " & _
          "WHERE Tipo = 'PALABRA' " & _
          "AND Idioma = '" & idioma & "' " & _
          "AND Palabra = '" & UCase(palabra) & "' " & _
          "AND Activo = True"

    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        BuscarExcepcionPalabra = rs!FonemaCompleto
    Else
        BuscarExcepcionPalabra = ""
    End If

    rs.Close
    Set rs = Nothing

End Function


