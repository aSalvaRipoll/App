Option Compare Database
Option Explicit

' ============================================================
' UDTs ICO estándar
' ============================================================

Private Type ICONDIR
    idReserved As Integer
    idType As Integer
    idCount As Integer
End Type

Private Type ICONDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer
    dwBytesInRes As Long
    dwImageOffset As Long
End Type

' ============================================================
' API
' ============================================================

Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare PtrSafe Function LoadImage Lib "user32" Alias "LoadImageA" _
        (ByVal hInst As LongPtr, ByVal lpsz As String, _
         ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, _
         ByVal un2 As Long) As LongPtr

Private Const IMAGE_ICON = 1
Private Const LR_LOADFROMFILE = &H10

' ============================================================
' Obtener stdPicture desde ImageMso
' ============================================================

Public Function GetPictureFromMso(pMso As String, Optional pSize As Long = 32) As stdPicture
    On Error Resume Next
    Set GetPictureFromMso = Application.CommandBars.GetImageMso(pMso, pSize, pSize)
End Function

' ============================================================
' Convertir BMP --> ICO (versión blindada)
' ============================================================

Public Function ConvertBMPtoICO(BmpPath As String, IcoPath As String) As Boolean
    On Error GoTo ErrHandler

    Dim bmp() As Byte
    Dim ico() As Byte
    Dim f As Integer
    Dim bmpSize As Long

    ' Leer BMP completo
    f = FreeFile
    Open BmpPath For Binary As #f
        bmpSize = LOF(f)
        ReDim bmp(0 To bmpSize - 1)
        Get #f, , bmp
    Close #f

    ' ------------------------------------------------------------
    ' 1. Leer cabecera BMP
    ' ------------------------------------------------------------

    Dim width As Long
    Dim height As Long
    Dim bitCount As Integer
    Dim dataOffset As Long
    Dim dibSize As Long

    CopyMemory width, bmp(18), 4
    CopyMemory height, bmp(22), 4
    CopyMemory bitCount, bmp(28), 2
    CopyMemory dataOffset, bmp(10), 4

    dibSize = bmpSize - 14

    ' ------------------------------------------------------------
    ' 2. Preparar cabecera ICO
    ' ------------------------------------------------------------

    Dim header As ICONDIR
    header.idReserved = 0
    header.idType = 1
    header.idCount = 1

    Dim entry As ICONDIRENTRY
    entry.bWidth = width
    entry.bHeight = height \ 2
    entry.bColorCount = 0
    entry.bReserved = 0
    entry.wPlanes = 1
    entry.wBitCount = bitCount
    entry.dwBytesInRes = dibSize
    entry.dwImageOffset = Len(header) + Len(entry)

    ' ------------------------------------------------------------
    ' 3. Construir ICO completo
    ' ------------------------------------------------------------

    ReDim ico(0 To entry.dwImageOffset + entry.dwBytesInRes - 1)

    CopyMemory ico(0), header, Len(header)
    CopyMemory ico(Len(header)), entry, Len(entry)
    CopyMemory ico(entry.dwImageOffset), bmp(14), dibSize

    ' ------------------------------------------------------------
    ' 4. Guardar ICO
    ' ------------------------------------------------------------

    f = FreeFile
    Open IcoPath For Binary As #f
        Put #f, , ico
    Close #f

    ConvertBMPtoICO = True
    Exit Function

ErrHandler:
    ConvertBMPtoICO = False
End Function

' ============================================================
' Obtener HICON desde ImageMso (blindado)
' ============================================================

Public Function GetIconFromImageMso_Blindado(ImageMso As String, TempName As String, _
                                             Optional Size As Long = 32) As LongPtr
    On Error GoTo ErrHandler

    Dim pic As stdPicture
    Dim bmpPath As String
    Dim icoPath As String

    bmpPath = Environ$("TEMP") & "\" & TempName & ".bmp"
    icoPath = Environ$("TEMP") & "\" & TempName & ".ico"

    ' Obtener imagen MSO
    Set pic = GetPictureFromMso(ImageMso, Size)
    If pic Is Nothing Then Exit Function

    ' Guardar BMP
    SavePicture pic, bmpPath

    ' Convertir a ICO blindado
    If ConvertBMPtoICO(bmpPath, icoPath) Then
        GetIconFromImageMso_Blindado = LoadImage(0, icoPath, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    End If

    Exit Function

ErrHandler:
    GetIconFromImageMso_Blindado = 0
End Function

' ============================================================
' Resolver icono (archivo o ImageMso)
' ============================================================

Public Function ResolveIconFile_Blindado(Icon As String, TempName As String) As LongPtr

    If Len(Icon) = 0 Then Exit Function

    If Left$(Icon, 9) = "ImageMso:" Then
        ResolveIconFile_Blindado = GetIconFromImageMso_Blindado(Mid$(Icon, 10), TempName)
    ElseIf Len(Dir$(Icon)) > 0 Then
        ResolveIconFile_Blindado = LoadImage(0, Icon, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
    End If

End Function
