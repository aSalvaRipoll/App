Attribute VB_Name = "modMotorFonetico_V2_1_DTO"
Option Compare Database
Option Explicit


' ============================================================
'   TestFonemas_ES
'   Prueba completa del motor fonético por sílaba (KOSMOS 2.1)
' ============================================================

Public Sub TestFonemas_ES()

    Dim res As Collection
    Dim sil As Variant
    Dim f As Variant

    Dim texto As String
    texto = "MARÍA DE LAS VIRTUDES"

    Debug.Print "Probando: "; texto
    Debug.Print String(40, "-")
    
    texto = LCase$(texto)

    ' Ejecutar el conversor principal
'    Set res = ConvertirTextoAFonemas_ES(texto, False)

    ' Recorrer la estructura devuelta
    For Each sil In res

        If IsObject(sil) Then
            ' Es una sílaba ? imprimir fonemas
            Debug.Print "Sílaba -> ";
            For Each f In sil
                Debug.Print f & " ";
            Next
            Debug.Print
        Else
            ' Es un separador entre palabras
            Debug.Print "---- SEPARADOR ----"
        End If

    Next

    Debug.Print String(40, "-")
    Debug.Print "FIN DE PRUEBA"

End Sub





'Sub TestFonemas()
'
'Dim col As Collection
'Dim s As String
'
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("María José", "es")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Col·lell", "ca")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Aitana", "ca-va")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Raül", "ca-ib")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Xoán", "gl")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Ainhoa", "eu")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("João", "pt-br")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Gonçalo", "pt-eu")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("Maël", "fr")
'    Set col = MF21_ConvertirNombreAParGrafemaIDFonema("McArthur", "en-gb")
'
'    s = FonemasEnColumnas(col)
'
'Debug.Print s
'
'End Sub

Sub TestFonemas()

    Debug.Print "=== CASTELLANO ==="
    ImprimirFonemas "María José", "es"

    Debug.Print "=== CATALÁN ==="
    ImprimirFonemas "Col·lell", "ca"

    Debug.Print "=== VALENCIANO ==="
    ImprimirFonemas "Aitana", "ca-va"

    Debug.Print "=== MALLORQUÍN ==="
    ImprimirFonemas "Raül", "ca-ib"

    Debug.Print "=== GALLEGO ==="
    ImprimirFonemas "Xoán", "gl"

    Debug.Print "=== EUSKARA ==="
    ImprimirFonemas "Ainhoa", "eu"

    Debug.Print "=== PORTUGUÉS BR ==="
    ImprimirFonemas "João", "pt-br"

    Debug.Print "=== PORTUGUÉS EU ==="
    ImprimirFonemas "Gonçalo", "pt-eu"

    Debug.Print "=== FRANCÉS ==="
    ImprimirFonemas "Maël", "fr"

    Debug.Print "=== INGLÉS GB ==="
    ImprimirFonemas "McArthur", "en-gb"

End Sub

Private Sub ImprimirFonemas(ByVal nombre As String, ByVal idioma As String)
    Dim col As Collection
    Dim s As String

    Set col = MF21_ConvertirNombreAParGrafemaIDFonema(nombre, idioma)
    s = FonemasEnColumnas(col)

    Debug.Print nombre & "  [" & idioma & "]"
    Debug.Print s
    Debug.Print
End Sub


Public Function FonemasEnColumnas(col As Collection) As String
    Dim f As clsFonema
    Dim item As Variant
    Dim maxGraf As Long, maxId As Long, maxVal As Long, maxVoc As Long
    Dim linea As String
    Dim salida As String
    
    ' --------------------------------------------------------
    ' 1. Calcular anchos máximos
    ' --------------------------------------------------------
    For Each item In col
        Set f = item
        
        If Len(f.GrafemaOri) > maxGraf Then maxGraf = Len(f.GrafemaOri)
        If Len(CStr(f.idFonema)) > maxId Then maxId = Len(CStr(f.idFonema))
        If Len(CStr(f.valor)) > maxVal Then maxVal = Len(CStr(f.valor))
        If Len(CStr(f.esVocal)) > maxVoc Then maxVoc = Len(CStr(f.esVocal))
    Next item
    
    ' --------------------------------------------------------
    ' 2. Cabecera
    ' --------------------------------------------------------
    salida = _
        Pad("GRAF", maxGraf) & "  " & _
        Pad("ID", maxId) & "  " & _
        Pad("VAL", maxVal) & "  " & _
        Pad("VOC", maxVoc) & vbCrLf
    
    salida = salida & String(Len(salida), "-") & vbCrLf
    
    ' --------------------------------------------------------
    ' 3. Filas
    ' --------------------------------------------------------
    For Each item In col
        Set f = item
        
        linea = _
            Pad(f.GrafemaOri, maxGraf) & "  " & _
            Pad(CStr(f.idFonema), maxId) & "  " & _
            Pad(CStr(f.valor), maxVal) & "  " & _
            Pad(CStr(f.esVocal), maxVoc)
        
        salida = salida & linea & vbCrLf
    Next item
    
    FonemasEnColumnas = salida
End Function

Private Function Pad(txt As String, ancho As Long) As String
    Pad = txt & Space$(ancho - Len(txt))
End Function


