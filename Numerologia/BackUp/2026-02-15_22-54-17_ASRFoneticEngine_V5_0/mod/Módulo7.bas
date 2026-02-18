Attribute VB_Name = "Módulo7"

'Option Compare Database
'Option Explicit
'
'
'Private modoDebug As Boolean   ' Activa / desactiva tracking
'
'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim cadenaNucleos As String
'    Dim nucleos() As String
'    Dim silabas() As String
'    Dim i As Byte, total As Byte
'
'    modoDebug = True
'
'    If modoDebug Then
'        addLog
'        addLog "============================================================="
'        addLog "            LOG DE DEPURACIÓN DE RESULTADOS "
'        addLog "============================================================="
'        addLog
'        addLog "Entrada: '" & Texto & "'"
'    End If
'
'    Texto = LCase$(Texto)
'    If modoDebug Then
'        addLog "Normalizado: '" & Texto & "'"
'        addLog "Total letras: " & Len(Texto)
'    End If
'
'    ' 1. Localizar núcleos ? cadena "3,7,10..."
'    cadenaNucleos = LocalizarNucleos(Texto)
'
'    If modoDebug Then
'        addLog
'        addLog "cadenaNucleos: " & cadenaNucleos
'    End If
'
'    ' 2. Convertir a matriz segura
'    nucleos = Split(cadenaNucleos, ",")
'
'    total = UBound(nucleos) + 1
'    ReDim silabas(1 To total)
'
'    ' 3. Construir sílabas
'    For i = 1 To total
'        silabas(i) = ConstruirSilaba(Texto, cadenaNucleos, i)
'    Next i
'
'    ' 4. Unir con separador
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
'Private Function LocalizarNucleos(ByVal T As String) As String
'    Dim i As Byte, L As Byte
'    Dim c As String
'    Dim lista As String
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento LocalizarNucleos"
'    End If
'
'    L = Len(T)
'
'    For i = 1 To L
'        c = Mid$(T, i, 1)
'        If EsVocal(c) Then
'            If lista <> "" Then lista = lista & ","
'            lista = lista & CStr(i)
'        End If
'    Next i
'
'    LocalizarNucleos = lista
'
'    If modoDebug Then
'        addLog "LocalizarNucleos: " & LocalizarNucleos
'        addLog " Fin LocalizarNucleos"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function ConstruirSilaba(ByVal T As String, ByVal cadenaN As String, ByVal idx As Byte) As String
'    Dim N() As String
'    Dim ataque As String, nucleo As String, coda As String
'    Dim pos As Byte
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento ConstruirSilaba"
'    End If
'
'    N = Split(cadenaN, ",")
'
'    pos = CByte(N(idx - 1))   ' N() base 0
'    nucleo = Mid$(T, pos, 1)
'
'    If modoDebug Then
'        addLog
'        addLog "T: " & T
'        addLog "cadenaN: " & cadenaN
'        addLog "idx: " & idx
'    End If
'
'    ataque = ObtenerAtaque(T, cadenaN, idx)
'    coda = ObtenerCoda(T, cadenaN, idx)
'
'    If modoDebug Then
'        addLog ""
'        addLog "ataque: " & ataque
'        addLog "nucleo: " & nucleo
'        addLog "coda: " & coda
'    End If
'
'    ConstruirSilaba = ataque & nucleo & coda
'
'    If modoDebug Then
'        addLog
'        addLog "ConstruirSilaba: " & ConstruirSilaba
'        addLog
'        addLog " Fin ConstruirSilaba"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function ObtenerAtaque(ByVal T As String, ByVal cadenaN As String, ByVal idx As Byte) As String
'    Dim N() As String
'    Dim inicio As Byte, fin As Byte
'    Dim grupo As String
'    Dim pos As Byte
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento ObtenerAtaque"
'    End If
'
'    N = Split(cadenaN, ",")
'
'    pos = CByte(N(idx - 1))
'
'    If modoDebug Then addLog "idx: " & idx
'
'    If idx = 1 Then
'        inicio = 1
'    Else
'        inicio = CByte(N(idx - 2)) + 1
'    End If
'
'    If modoDebug Then addLog "inicio: " & inicio
'
'    fin = pos - 1
'
'    If modoDebug Then addLog "fin: " & fin
'
'    If fin < inicio Then Exit Function
'
'    grupo = Mid$(T, inicio, fin - inicio + 1)
'
'    If modoDebug Then addLog "grupo: " & grupo
'
'    ObtenerAtaque = ResolverAtaque(grupo)
'
'    If modoDebug Then
'        addLog "ObtenerAtaque: " & ObtenerAtaque
'        addLog
'        addLog " Fin ObtenerAtaque"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function ResolverAtaque(ByVal g As String) As String
'    Dim L As Byte: L = Len(g)
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento ResolverAtaque"
'    End If
'
'    If L = 0 Then Exit Function
'    If L = 1 Then ResolverAtaque = g: Exit Function
'
'    If L = 2 Then
'        If EsAtaqueComplejo(g) Then ResolverAtaque = g Else ResolverAtaque = Right$(g, 1)
'        Exit Function
'    End If
'
'    If L = 3 Then
'        If EsAtaqueComplejo(Right$(g, 2)) Then
'            ResolverAtaque = Right$(g, 2)
'        Else
'            ResolverAtaque = Right$(g, 1)
'        End If
'        Exit Function
'    End If
'
'    If L >= 4 Then
'        ResolverAtaque = Right$(g, 2)
'        Exit Function
'    End If
'
'    If modoDebug Then
'        addLog
'        addLog " Fin ResolverAtaque"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function ObtenerCoda(ByVal T As String, ByVal cadenaN As String, ByVal idx As Byte) As String
'    Dim N() As String
'    Dim inicio As Byte, fin As Byte
'    Dim grupo As String
'    Dim pos As Byte
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento ObtenerCoda"
'    End If
'
'    N = Split(cadenaN, ",")
'
'    pos = CByte(N(idx - 1))
'
'    inicio = pos + 1
'
'    If idx = UBound(N) + 1 Then
'        fin = Len(T)
'    Else
'        fin = CByte(N(idx)) - 1
'    End If
'
'    If fin < inicio Then Exit Function
'
'    grupo = Mid$(T, inicio, fin - inicio + 1)
'
'    ObtenerCoda = ResolverCoda(grupo)
'
'    If modoDebug Then
'        addLog
'        addLog " Fin ObtenerCoda"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function ResolverCoda(ByVal g As String) As String
'    Dim L As Byte: L = Len(g)
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento ResolverCoda"
'    End If
'
'    If L = 0 Then Exit Function
'
'    ' *** CAMBIO CLAVE ***
'    ' 1 consonante entre vocales ? NO coda, va al ataque siguiente
'    If L = 1 Then Exit Function
'
'    If L = 2 Then
'        If EsCodaCompleja(g) Then ResolverCoda = g Else ResolverCoda = Left$(g, 1)
'        Exit Function
'    End If
'
'    If L = 3 Then
'        If Right$(g, 1) = "s" Then ResolverCoda = Left$(g, 2) Else ResolverCoda = Left$(g, 1)
'        Exit Function
'    End If
'
'    If L >= 4 Then
'        ResolverCoda = Left$(g, 2)
'        Exit Function
'    End If
'
'    If modoDebug Then
'        addLog
'        addLog " Fin ResolverCoda"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function EsVocal(ByVal c As String) As Boolean
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento EsVocal"
'    End If
'
'    EsVocal = InStr("aeiouáéíóúü", c) > 0
'
'    If modoDebug Then
'        addLog "c: " & c & " --> EsVocal: " & EsVocal
'        addLog
'        addLog " Fin EsVocal"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function EsAtaqueComplejo(ByVal g As String) As Boolean
'    Dim AC As Variant
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento EsAtaqueComplejo"
'    End If
'
'    AC = Array("pr", "br", "tr", "dr", "cr", "kr", "gr", "fr", _
'               "pl", "bl", "cl", "kl", "gl", "fl", "tl")
'
'    EsAtaqueComplejo = (UBound(Filter(AC, g)) >= 0)
'
'    If modoDebug Then
'        addLog "g: " & g & " --> EsAtaqueComplejo: " & EsAtaqueComplejo
'        addLog
'        addLog " Fin EsAtaqueComplejo"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'Private Function EsCodaCompleja(ByVal g As String) As Boolean
'    Dim CC As Variant
'
'    If modoDebug Then
'        addLog
'        addLog "---------------------------------------"
'        addLog " Procedimiento EsCodaCompleja"
'    End If
'
'    CC = Array("ns", "rts", "rps", "lts", "nts", "mps")
'
'    EsCodaCompleja = (UBound(Filter(CC, g)) >= 0)
'
'    If modoDebug Then
'        addLog "g: " & g & " --> EsCodaCompleja: " & EsCodaCompleja
'        addLog
'        addLog " Fin EsCodaCompleja"
'        addLog "---------------------------------------"
'    End If
'End Function
'
'
