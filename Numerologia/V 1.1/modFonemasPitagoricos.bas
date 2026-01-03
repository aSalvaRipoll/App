Attribute VB_Name = "modFonemasPitagoricos"

Option Compare Database
Option Explicit

' ---------------------------------------------------------
'  Coleccion de fonemas cargados desde la tabla maestra
'  colFonemas("A") = 1
'  colFonemas("CH") = 3
' ---------------------------------------------------------
Public colFonemas As Collection

Public Sub CargarFonemasPitagoricos()
    Dim rs As DAO.Recordset
    Dim clave As String
    Dim valor As Byte
    
    Set colFonemas = New Collection
    
    Set rs = CurrentDb.OpenRecordset( _
        "SELECT FonemaASCII, Valor FROM tbmFonemas ORDER BY Longitud DESC")
    
    Do While Not rs.EOF
        clave = UCase$(Nz(rs!FonemaASCII, ""))
        valor = Nz(rs!valor, 0)
        
        If clave <> "" Then
            On Error Resume Next
            colFonemas.Add valor, clave
            On Error GoTo 0
        End If
        
        rs.MoveNext
    Loop
    
    rs.Close
    Set rs = Nothing
End Sub

Public Function ExisteFonema(ByVal f As String) As Boolean
    On Error GoTo ErrHandler
    Dim tmp As Variant
    
    f = UCase$(f)
    tmp = colFonemas(f)
    
    ExisteFonema = True
    Exit Function
    
ErrHandler:
    ExisteFonema = False
End Function

Public Function ValorFonema(ByVal f As String) As Byte
    On Error GoTo ErrHandler
    
    f = UCase$(f)
    ValorFonema = colFonemas(f)
    Exit Function
    
ErrHandler:
    ValorFonema = 0
End Function



'Option Compare Database
'Option Explicit
'
'' ---------------------------------------------------------
''  Suma total de valores del array arrFonemas()
'' ---------------------------------------------------------
'Public Function ValorTotalFonemas() As Long
'    Dim i As Long
'    Dim total As Long
'
'    For i = LBound(arrFonemas) To UBound(arrFonemas)
'        total = total + arrFonemas(i).valor
'    Next i
'
'    ValorTotalFonemas = total
'End Function
'
'' ---------------------------------------------------------
''  Reduccion pitagorica (ej: 27 --> 9)
'' ---------------------------------------------------------
'Public Function ReduccionPitagorica(ByVal valor As Long) As Long
'    Do While valor > 9
'        valor = SumaDigitos(valor)
'    Loop
'    ReduccionPitagorica = valor
'End Function
'
'Private Function SumaDigitos(ByVal n As Long) As Long
'    Dim s As Long
'    Do While n > 0
'        s = s + (n Mod 10)
'        n = n \ 10
'    Loop
'    SumaDigitos = s
'End Function



'Option Compare Database
'Option Explicit
'
'Public Fonemas As Collection
'
'Public Sub CargarFonemasPitagoricos()
'    Dim rs As DAO.Recordset
'    Set Fonemas = New Collection
'
'    Set rs = CurrentDb.OpenRecordset("SELECT Fonema, ValorPitagorico FROM tbmFonemas")
'
'    Do While Not rs.EOF
'        Fonemas.Add rs!ValorPitagorico, UCase(rs!fonema)
'        rs.MoveNext
'    Loop
'
'    rs.Close
'    Set rs = Nothing
'End Sub
'
'Public Fonemas As Collection
'
'Public Function ValorPitagoricoDeClave(ByVal clave As String) As Long
'    Dim i As Long
'    Dim token As String
'    Dim total As Long
'
'    clave = UCase(clave)
'    i = 1
'
'    Do While i <= Len(clave)
'
'        ' Intentar fonemas dobles (SH, NY, LY, CH)
'        If i < Len(clave) Then
'            token = Mid(clave, i, 2)
'            If ExisteFonema(token) Then
'                total = total + Fonemas(token)
'                i = i + 2
'                GoTo siguiente
'            End If
'        End If
'
'        ' Intentar fonema simple
'        token = Mid(clave, i, 1)
'        If ExisteFonema(token) Then
'            total = total + Fonemas(token)
'        End If
'
'siguiente:
'        i = i + 1
'    Loop
'
'    ValorPitagoricoDeClave = total
'End Function
'
'Private Function ExisteFonema(f As String) As Boolean
'    On Error GoTo ErrHandler
'    Dim tmp As Variant
'    tmp = Fonemas(f)
'    ExisteFonema = True
'    Exit Function
'ErrHandler:
'    ExisteFonema = False
'End Function
'
