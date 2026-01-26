Attribute VB_Name = "modMotorFonetico_V2_1"

Option Compare Database
Option Explicit

' ============================================================
'   MÓDULO PRINCIPAL — Motor Fonético 2.1 (versión KOSMOS)
'   - Tokeniza grafema por grafema
'   - Aplica reglas por idioma
'   - Genera pares GrafemaOriginal – idFonema
' ============================================================

Public Function MF21_ConvertirNombreAParGrafemaIDFonema( _
        ByVal NombreOriginal As String, _
        ByVal Abreviado As String _
    ) As Collection

    Dim texto As String
    Dim col As New Collection
    Dim i As Long
    Dim graf As String
    Dim idF As Byte
    Dim NumOrden As Long
    Dim ant As String, sig As String

    Dim exPal As String
    
    ' --------------------------------------------------------
    ' 0. Cargar colección
    ' --------------------------------------------------------
    
    If colFonemas Is Nothing Then CargarFonemas

    ' --------------------------------------------------------
    ' 1. Normalización previa
    ' --------------------------------------------------------
    texto = Trim$(NombreOriginal)
    
    ' Eliminar espacios dobles
    While InStr(texto, "  ")
        texto = Replace(texto, "  ", " ")
    Wend
    
    ' Si está vacío, no seguimos
    If Len(texto) = 0 Then
        Set MF21_ConvertirNombreAParGrafemaIDFonema = col
        Exit Function
    End If
    
    ' Convertir a mayúsculas
    texto = UCase$(texto)
    
    ' Normalizar vocales según idioma
    texto = MF_NormalizarVocalesPorIdioma(texto, Abreviado)

    ' --------------------------------------------------------
    ' 1.b Excepción por palabra completa
    ' --------------------------------------------------------
    
'    exPal = BuscarExcepcionPalabra(texto, Abreviado)
'
'    If exPal <> "" Then
'        ' Construir colección directamente desde el fonema completo
'        Dim colEx As New Collection
'        Dim j As Long, c As String
'
'        For j = 1 To Len(exPal)
'            c = Mid$(exPal, j, 1)
'            Dim idFex As Byte
'            idFex = AscW(c) ' o tu método para mapear fonema ? idFonema
'
'            If idFex <> 0 Then
'                RegistrarFonema colEx, c, idFex, j
'            End If
'        Next j
'
'        Set MF21_ConvertirNombreAParGrafemaIDFonema = colEx
'        Exit Function
'    End If


    ' --------------------------------------------------------
    ' 2. Tokenización grafema por grafema
    ' --------------------------------------------------------
    i = 1
    NumOrden = 1

    Do While i <= Len(texto)

        ' Obtener contexto
'        ant = IIf(i > 1, Mid$(texto, i - 1, 1), "")
'        sig = IIf(i < Len(texto), Mid$(texto, i + 1, 1), "")
        ' Obtener contexto seguro
        If i > 1 Then
            ant = Mid$(texto, i - 1, 1)
        Else
            ant = ""
        End If
        
        If i < Len(texto) Then
            sig = Mid$(texto, i + 1, 1)
        Else
            sig = ""
        End If

        ' Intentar trigrafema
        If i <= Len(texto) - 2 Then
            graf = Mid$(texto, i, 3)
            idF = ConvertirGrafemaAIdFonema(graf, Abreviado, ant, sig)
            If idF <> 0 Then
                RegistrarFonema col, graf, idF, NumOrden
                NumOrden = NumOrden + 1
                i = i + 3
                GoTo Siguiente
            End If
        End If

        ' Intentar dígrafo
        If i <= Len(texto) - 1 Then
            graf = Mid$(texto, i, 2)
            idF = ConvertirGrafemaAIdFonema(graf, Abreviado, ant, sig)
            If idF <> 0 Then
                RegistrarFonema col, graf, idF, NumOrden
                NumOrden = NumOrden + 1
                i = i + 2
                GoTo Siguiente
            End If
        End If

        ' Monógrafo
        graf = Mid$(texto, i, 1)
        idF = ConvertirGrafemaAIdFonema(graf, Abreviado, ant, sig)

        If idF <> 0 Then
            RegistrarFonema col, graf, idF, NumOrden
            NumOrden = NumOrden + 1
        End If

        i = i + 1

Siguiente:
    Loop

    Set MF21_ConvertirNombreAParGrafemaIDFonema = col
    
    Dim objfon As clsFonema
    For Each objfon In col
        Debug.Print objfon.NumOrden; " - "; objfon.GrafemaOri; " > "; objfon.ASCII
    Next
    
End Function

Private Sub RegistrarFonema( _
        ByRef col As Collection, _
        ByVal graf As String, _
        ByVal idFonema As Byte, _
        ByVal NumOrden As Long _
    )

    Dim f As New clsFonema
    Dim info As clsFonema
    
    
     Set info = GetFonema(idFonema)
     
    ' Datos básicos
    f.GrafemaOri = graf
    f.idFonema = idFonema
    f.NumOrden = NumOrden

     'f.fonema = info.fonema
    f.EsVocal = info.EsVocal
    f.Valor = info.Valor

    col.Add f

End Sub

Public Function GetFonema(ByVal id As Byte) As clsFonema
    Set GetFonema = colFonemas(CStr(id))
End Function


Public Function ConvertirGrafemaAIdFonema( _
        ByVal graf As String, _
        ByVal idioma As String, _
        ByVal ant As String, _
        ByVal sig As String _
    ) As Byte

Dim ex As Byte

' ============================================================
' 0. Excepciones por grafema
' ============================================================

'    ex = BuscarExcepcion(graf, idioma)
'
'    If ex <> 0 Then
'        ConvertirGrafemaAIdFonema = ex
'        Exit Function
'    End If


    Select Case LCase$(idioma)

        Case "es":      ConvertirGrafemaAIdFonema = ReglasCastellano(graf, ant, sig, False)
        Case "ca":      ConvertirGrafemaAIdFonema = ReglasCatala(graf, ant, sig, False)
        Case "ca-ib":   ConvertirGrafemaAIdFonema = ReglasMallorquin(graf, ant, sig, False)
        Case "ca-va":   ConvertirGrafemaAIdFonema = ReglasValenciano(graf, ant, sig, False)
        Case "eu":      ConvertirGrafemaAIdFonema = ReglasEuskera(graf, ant, sig, False)
        Case "gl":      ConvertirGrafemaAIdFonema = ReglasGalego(graf, ant, sig, False)
        
        Case "pt":      ConvertirGrafemaAIdFonema = ReglasPortugues_PT_EU(graf, ant, sig, False)
        Case "pt-eu":   ConvertirGrafemaAIdFonema = ReglasPortugues_PT_EU(graf, ant, sig, False)
        Case "pt-br":   ConvertirGrafemaAIdFonema = ReglasPortugues_PT_BR(graf, ant, sig, False)

        Case "fr":      ConvertirGrafemaAIdFonema = ReglasFrances(graf, ant, sig, False)
        Case "en-gb":   ConvertirGrafemaAIdFonema = ReglasIngles(graf, ant, sig, False)
'        Case "en-us":   ConvertirGrafemaAIdFonema = ReglasIngles_EN_US(graf, ant, sig, False)
'        Case "en-us-af": ConvertirGrafemaAIdFonema = ReglasIngles_EN_US_AF(graf, ant, sig, False)
        Case Else:      ConvertirGrafemaAIdFonema = 0
    End Select

End Function

Public Function ObtenerFonemas( _
        ByVal texto As String, _
        ByVal idioma As eIdiomaFonetico _
    ) As Collection

    Dim col As New Collection
    Dim i As Long
    Dim graf As String, ant As String, sig As String
    Dim idF As Byte
    Dim lenTexto As Long
    
    texto = UCase$(texto)
    lenTexto = Len(texto)
    i = 1
    
    Do While i <= lenTexto
        
        ' 1) Obtener el siguiente grafema (trígrafo / dígrafo / monógrafo)
        graf = SiguienteGrafema(texto, i)
        
        ' 2) Contexto anterior y siguiente (un carácter, simplificado)
        If i > 1 Then
            ant = Mid$(texto, i - 1, 1)
        Else
            ant = ""
        End If
        
        If i + Len(graf) <= lenTexto Then
            sig = Mid$(texto, i + Len(graf), 1)
        Else
            sig = ""
        End If
        
        ' 3) Llamar al motor de reglas según idioma
        Select Case idioma
            Case idCastellano:  idF = ReglasCastellano(graf, ant, sig, False)
            Case idCatala:      idF = ReglasCatala(graf, ant, sig, False)
            Case idMallorquin:  idF = ReglasMallorquin(graf, ant, sig, False)
            Case idValenciano:  idF = ReglasValenciano(graf, ant, sig, False)
            Case idEuskera:     idF = ReglasEuskera(graf, ant, sig, False)
            Case idGalego:      idF = ReglasGalego(graf, ant, sig, False)
            
            Case idPortugues:   idF = ReglasPortugues(graf, ant, sig, False)
            Case idPortuguesEU: idF = ReglasPortugues_PT_EU(graf, ant, sig, False)
            Case idPortuguesBR: idF = ReglasPortugues_PT_BR(graf, ant, sig, False)

            Case idFrances:     idF = ReglasFrances(graf, ant, sig, False)
            Case idIngles:      idF = ReglasIngles(graf, ant, sig, False)
        End Select
        
        ' 4) Si hay fonema, lo añadimos
        If idF <> 0 Then
            col.Add idF
        End If
        
        ' 5) Avanzar índice según longitud del grafema
        i = i + Len(graf)
        
    Loop
    
    Set ObtenerFonemas = col

End Function

Private Function NG(g As String) As String

' ============================================================
'   FUNCIÓN LOCAL: normaliza grafemas combinados a precompuestos
' ============================================================

    ' Nasales
    g = Replace(g, "A" & ChrW(771), "Ã")
    g = Replace(g, "O" & ChrW(771), "Õ")

    ' Agudas
    g = Replace(g, "A" & ChrW(769), "Á")
    g = Replace(g, "E" & ChrW(769), "É")
    g = Replace(g, "I" & ChrW(769), "Í")
    g = Replace(g, "O" & ChrW(769), "Ó")
    g = Replace(g, "U" & ChrW(769), "Ú")

    ' Graves
    g = Replace(g, "A" & ChrW(768), "À")
    g = Replace(g, "E" & ChrW(768), "È")
    g = Replace(g, "I" & ChrW(768), "Ì")
    g = Replace(g, "O" & ChrW(768), "Ò")
    g = Replace(g, "U" & ChrW(768), "Ù")

    ' Circunflejos
    g = Replace(g, "A" & ChrW(770), "Â")
    g = Replace(g, "E" & ChrW(770), "Ê")
    g = Replace(g, "I" & ChrW(770), "Î")
    g = Replace(g, "O" & ChrW(770), "Ô")
    g = Replace(g, "U" & ChrW(770), "Û")

    ' Diéresis
    g = Replace(g, "A" & ChrW(776), "Ä")
    g = Replace(g, "E" & ChrW(776), "Ë")
    g = Replace(g, "I" & ChrW(776), "Ï")
    g = Replace(g, "O" & ChrW(776), "Ö")
    g = Replace(g, "U" & ChrW(776), "Ü")

    NG = g

End Function


Private Function SiguienteGrafema( _
        ByVal texto As String, _
        ByVal pos As Long _
    ) As String

'============================================
'=                                          =
'=            Versión KOSMOS                =
'=                                          =
'============================================
    Dim t As String
    Dim tri As String, di As String, mo As String
    Dim lenTexto As Long

    t = texto
    lenTexto = Len(t)

    ' ============================================================
    '   1. TRÍGRAFOS (incluye diacríticos combinados)
    ' ============================================================
    If pos + 2 <= lenTexto Then
        tri = NG(Mid$(t, pos, 3))

        Select Case tri

            ' Portugueses
            Case "GÜE", "GÜI", "GUE", "GUI", "QUE", "QUI"
                SiguienteGrafema = tri: Exit Function

            ' Nasales portuguesas/francesas
            Case "ÃO", "ÃE", "ÃI", "ÕE", "ÕI"
                SiguienteGrafema = tri: Exit Function

            ' Franceses
            Case "AIN", "EIN", "EIM", "OIN"
                SiguienteGrafema = tri: Exit Function

            ' Diptongos con acento
            Case "ÁI", "ÉI", "ÓI", "ÂI", "ÊI", "ÔI"
                SiguienteGrafema = tri: Exit Function

        End Select
    End If


    ' ============================================================
    '   2. DÍGRAFOS (incluye diacríticos combinados)
    ' ============================================================
    If pos + 1 <= lenTexto Then
        di = NG(Mid$(t, pos, 2))

        Select Case di

            ' Universales
            Case "CH", "LL", "RR", "NY", "GN", "PH", "SH", "TH", "DH", "WH"
                SiguienteGrafema = di: Exit Function

            ' Catalán / Mallorquín
            Case "TX", "TS", "TZ"
                SiguienteGrafema = di: Exit Function

            ' Portugués
            Case "NH", "LH", "SS"
                SiguienteGrafema = di: Exit Function

            ' Ela geminada
            Case "L·", "L."
                SiguienteGrafema = di: Exit Function

            ' Nasales PT/FR
            Case "AN", "AM", "EN", "EM", "IN", "IM", "ON", "OM", "UN", "UM", "YN", "YM"
                SiguienteGrafema = di: Exit Function

            ' Diptongos universales
            Case "AI", "EI", "OI", "OU", "AU", "EU", "UI"
                SiguienteGrafema = di: Exit Function

            ' Diptongos con acento
            Case "ÁI", "ÉU", "ÓI", "ÂO", "ÊI", "ÔU"
                SiguienteGrafema = di: Exit Function

        End Select
    End If


    ' ============================================================
    '   3. MONÓGRAFO (con diacríticos normalizados)
    ' ============================================================
    mo = NG(Mid$(t, pos, 1))
    SiguienteGrafema = mo

End Function


