Attribute VB_Name = "Módulo5"
Option Compare Database
Option Explicit

'
'
'Public Function SilabearPalabra(ByVal Texto As String) As String
'    Dim nucleos() As Long
'    Dim silabas() As String
'    Dim i As Long, total As Long
'
'    Texto = LCase$(Texto)
'
'    ' 1. Localizar núcleos (vocales)
'    nucleos = LocalizarNucleos(Texto)
'    total = UBound(nucleos)
'
'    ReDim silabas(1 To total)
'
'    ' 2. Construir sílabas
'    For i = 1 To total
'        silabas(i) = ConstruirSilaba(Texto, nucleos, i)
'    Next i
'
'    ' 3. Unir con separador
'    SilabearPalabra = Join(silabas, " | ")
'End Function
'
'Private Function LocalizarNucleos(ByVal T As String) As Long()
'    Dim pos() As Long
'    Dim i As Long, L As Long, c As String
'    Dim count As Long: count = 0
'
'    L = Len(T)
'    ReDim pos(1 To L)
'
'    For i = 1 To L
'        c = Mid$(T, i, 1)
'        If EsVocal(c) Then
'            count = count + 1
'            pos(count) = i
'        End If
'    Next i
'
'    ReDim Preserve pos(1 To count)
'    LocalizarNucleos = pos
'End Function
'
'Private Function ConstruirSilaba(ByVal T As String, N() As Long, ByVal idx As Long) As String
'    Dim ataque As String, nucleo As String, coda As String
'
'    nucleo = Mid$(T, N(idx), 1)
'
'    ataque = ObtenerAtaque(T, N, idx)
'    coda = ObtenerCoda(T, N, idx)
'
'    ConstruirSilaba = ataque & nucleo & coda
'End Function
'
'Private Function ObtenerAtaque(ByVal T As String, N() As Long, ByVal idx As Long) As String
'    Dim inicio As Long, fin As Long
'    Dim grupo As String
'
'    inicio = IIf(idx = 1, 1, N(idx - 1) + 1)
'    fin = N(idx) - 1
'
'    If fin < inicio Then Exit Function
'
'    grupo = Mid$(T, inicio, fin - inicio + 1)
'
'    ObtenerAtaque = ResolverAtaque(grupo)
'End Function
'
'Private Function ResolverAtaque(ByVal g As String) As String
'    Dim L As Long: L = Len(g)
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
'End Function
'
'Private Function ObtenerCoda(ByVal T As String, N() As Long, ByVal idx As Long) As String
'    Dim inicio As Long, fin As Long
'    Dim grupo As String
'
'    inicio = N(idx) + 1
'    fin = IIf(idx = UBound(N), Len(T), N(idx + 1) - 1)
'
'    If fin < inicio Then Exit Function
'
'    grupo = Mid$(T, inicio, fin - inicio + 1)
'
'    ObtenerCoda = ResolverCoda(grupo)
'End Function
'
'Private Function ResolverCoda(ByVal g As String) As String
'    Dim L As Long: L = Len(g)
'
'    If L = 0 Then Exit Function
'    If L = 1 Then ResolverCoda = g: Exit Function
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
'End Function
'
'Private Function EsVocal(ByVal c As String) As Boolean
'    EsVocal = InStr("aeiouáéíóúü", c) > 0
'End Function
'
'Private Function EsAtaqueComplejo(ByVal g As String) As Boolean
'    Dim AC As Variant
'    AC = Array("pr", "br", "tr", "dr", "cr", "kr", "gr", "fr", _
'               "pl", "bl", "cl", "kl", "gl", "fl", "tl")
'
'    EsAtaqueComplejo = (UBound(Filter(AC, g)) >= 0)
'End Function
'
'Private Function EsCodaCompleja(ByVal g As String) As Boolean
'    Dim CC As Variant
'    CC = Array("ns", "rts", "rps", "lts", "nts", "mps")
'
'    EsCodaCompleja = (UBound(Filter(CC, g)) >= 0)
'End Function
'
'
