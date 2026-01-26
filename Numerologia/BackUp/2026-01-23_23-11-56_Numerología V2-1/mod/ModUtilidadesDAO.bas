Attribute VB_Name = "ModUtilidadesDAO"
' ------------------------------------------------------
' Nombre:    ModUtilidadesDAO
' Tipo:      Módulo
' Propósito:
' Autor:     asalv
' Fecha:     15/01/2026
' ------------------------------------------------------


Option Compare Database
Option Explicit

' ============================================================================
'   Módulo: ModUtilidadesDAO
'   Autor:  Alba Salvá
'   Año:    2026
'
'   Propósito:
'       Conjunto de funciones utilitarias para acceso a datos en Access.
'       Incluye funciones locales (prefijo s) y funciones extendidas/remotas
'       (prefijo r), con soporte para bases externas, JetLink y validación.
'
'   Notas:
'       - Todas las funciones usan DAO explícito.
'       - Limpieza segura de objetos.
'       - No se cierra CurrentDb nunca.
'       - SQL seguro con corchetes mediante SafeName.
' ============================================================================


' ============================================================================
'   Función auxiliar para entrecorchetar nombres de campos/tablas
' ============================================================================
Public Function SafeName(ByVal Nombre As String) As String
    SafeName = "[" & Replace(Nombre, "]", "]]") & "]"
End Function


' ============================================================================
'   FUNCIONES SENCILLAS (LOCAL ONLY)
' ============================================================================

' -----------------------------------------
' sMax – Máximo de un campo
' -----------------------------------------
Public Function sMax(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "") As Long
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT MAX(" & SafeName(Campo) & ") AS MiMax FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sMax = 0
    Else
        sMax = Nz(rs!MiMax, 0)
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    sMax = 0
    Resume Salir
End Function


' -----------------------------------------
' sCount – Cuenta de registros
' -----------------------------------------
Public Function sCount(ByVal Campo As String, _
                       ByVal Tabla As String, _
                       Optional ByVal MiWhere As String = "") As Long
                       
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim c1 As String
    
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    If Trim$(Campo) = "*" Then
        c1 = Campo
    Else
        c1 = SafeName(Campo)
    End If
    
    sql = "SELECT COUNT(" & c1 & ") AS Cuenta FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sCount = 0
    Else
        sCount = Nz(rs!Cuenta, 0)
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    sCount = 0
    Resume Salir
End Function

' -----------------------------------------
' sExists / sExistsFast – Devuelve True
'          si existe al menos 1 registro
' -----------------------------------------
Public Function sExists(ByVal Campo As String, _
                        ByVal Tabla As String, _
                        Optional ByVal MiWhere As String = "") As Boolean

    sExists = (sCount(Campo, Tabla, MiWhere) > 0)

End Function


Public Function sExistsFast(ByVal Tabla As String, _
                            Optional ByVal MiWhere As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String

    Set db = CurrentDb

    sql = "SELECT 1 FROM " & SafeName(Tabla)

    If Trim$(MiWhere) <> "" Then
        sql = sql & " WHERE " & Trim$(MiWhere)
    End If

    sql = sql & " LIMIT 1"   ' Truco para acelerar

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)

    sExistsFast = Not rs.EOF

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    sExistsFast = False
    Resume Salir
End Function

' -----------------------------------------
' sLookup – Devuelve un único valor
' -----------------------------------------
Public Function sLookup(ByVal Campo As String, _
                        ByVal Tabla As String, _
                        Optional ByVal MiWhere As String = "") As Variant
                        
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS Dato FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sLookup = Null
    Else
        sLookup = rs!Dato
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    sLookup = Null
    Resume Salir
End Function


' -----------------------------------------
' sFirst – Primer valor de un campo
' -----------------------------------------
Public Function sFirst(ByVal Campo As String, _
                       ByVal Tabla As String, _
                       Optional ByVal MiWhere As String = "") As Variant
                       
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS Dato FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " ORDER BY " & SafeName(Campo) & " ASC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sFirst = Null
    Else
        sFirst = rs!Dato
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    sFirst = Null
    Resume Salir
End Function


' -----------------------------------------
' sLast – Último valor de un campo
' -----------------------------------------
Public Function sLast(ByVal Campo As String, _
                      ByVal Tabla As String, _
                      Optional ByVal MiWhere As String = "") As Variant
                      
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS Dato FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " ORDER BY " & SafeName(Campo) & " DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sLast = Null
    Else
        sLast = rs!Dato
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    sLast = Null
    Resume Salir
End Function


' -----------------------------------------
' sAvg – Media aritmética
' -----------------------------------------
Public Function sAvg(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "") As Variant
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT AVG(" & SafeName(Campo) & ") AS MiValor FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sAvg = Null
    Else
        sAvg = Nz(rs!MiValor, Null)
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    sAvg = Null
    Resume Salir
End Function


' -----------------------------------------
' sMedian – Mediana
' -----------------------------------------
Public Function sMedian(ByVal Campo As String, _
                        ByVal Tabla As String, _
                        Optional ByVal MiWhere As String = "") As Variant
                        
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim n As Long
    Dim pos As Long
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    ' 1. Contar registros válidos
    sql = "SELECT COUNT(" & SafeName(Campo) & ") AS N FROM " & SafeName(Tabla)
    If Trim(MiWhere) <> "" Then sql = sql & " WHERE " & MiWhere
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    If rs.EOF Then GoTo NoData
    n = Nz(rs!n, 0)
    rs.Close
    
    If n = 0 Then GoTo NoData
    
    ' 2. Obtener la mediana
    If (n Mod 2) = 1 Then
        ' Impar ? valor central
        pos = (n + 1) \ 2
        sql = "SELECT " & SafeName(Campo) & " AS V FROM " & SafeName(Tabla)
        If Trim(MiWhere) <> "" Then sql = sql & " WHERE " & MiWhere
        sql = sql & " ORDER BY " & SafeName(Campo) & " ASC"
        
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        rs.Move pos - 1
        sMedian = rs!v
    Else
        ' Par ? media de los dos centrales
        Dim Pos1 As Long, Pos2 As Long
        Dim v1 As Variant, V2 As Variant
        
        Pos1 = n \ 2
        Pos2 = Pos1 + 1
        
        sql = "SELECT " & SafeName(Campo) & " AS V FROM " & SafeName(Tabla)
        If Trim(MiWhere) <> "" Then sql = sql & " WHERE " & MiWhere
        sql = sql & " ORDER BY " & SafeName(Campo) & " ASC"
        
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
        rs.Move Pos1 - 1
        v1 = rs!v
        rs.MoveNext
        V2 = rs!v
        
        sMedian = (v1 + V2) / 2
    End If
    
    GoTo Salir

NoData:
    sMedian = Null
    GoTo Salir

ErrHandler:
    sMedian = Null

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Function


' -----------------------------------------
' sMode – Moda
' -----------------------------------------
Public Function sMode(ByVal Campo As String, _
                      ByVal Tabla As String, _
                      Optional ByVal MiWhere As String = "") As Variant
                      
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & ", COUNT(*) AS Frecuencia " & _
          "FROM " & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " GROUP BY " & SafeName(Campo) & _
                " ORDER BY COUNT(*) DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        sMode = Null
    Else
        sMode = rs.Fields(0).Value
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function
    
ErrHandler:
    sMode = Null
    Resume Salir
End Function


' ============================================================================
'   FUNCIONES EXTENDIDAS (REMOTAS / ROBUSTAS)
' ============================================================================

Private Function GetDB(ByVal DbPath As String, ByVal UseJetLink As Boolean) As dao.Database
    If DbPath = "" Or (DbPath <> "" And UseJetLink) Then
        Set GetDB = CurrentDb
    Else
        Set GetDB = DBEngine.OpenDatabase(DbPath)
    End If
End Function


' -----------------------------------------
' rCount – Cuenta remota
' -----------------------------------------
Public Function rCount(ByVal Campo As String, _
                       ByVal Tabla As String, _
                       Optional ByVal MiWhere As String = "", _
                       Optional ByVal DbPath As String = "", _
                       Optional ByVal UseJetLink As Boolean = True) As Long
                       
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Long
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT COUNT(" & SafeName(Campo) & ") AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        Valor = 0
    Else
        Valor = Nz(rs!MiValor, 0)
    End If
    
    rCount = Valor
    GoTo Salir

ErrHandler:
    rCount = 0
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rLookup – Lookup remoto
' -----------------------------------------
Public Function rLookup(ByVal Campo As String, _
                        ByVal Tabla As String, _
                        Optional ByVal MiWhere As String = "", _
                        Optional ByVal DbPath As String = "", _
                        Optional ByVal UseJetLink As Boolean = True) As Variant
                        
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Variant
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        Valor = Null
    Else
        Valor = rs!MiValor
    End If
    
    rLookup = Valor
    GoTo Salir

ErrHandler:
    rLookup = Null
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rSum – Suma remota
' -----------------------------------------
Public Function rSum(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "", _
                     Optional ByVal DbPath As String = "", _
                     Optional ByVal UseJetLink As Boolean = True) As Variant
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Variant
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT SUM(" & SafeName(Campo) & ") AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        Valor = Null
    Else
        Valor = Nz(rs!MiValor, 0)
    End If
    
    rSum = Valor
    GoTo Salir

ErrHandler:
    rSum = Null
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rMax – Máximo remoto
' -----------------------------------------
Public Function rMax(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "", _
                     Optional ByVal DbPath As String = "", _
                     Optional ByVal UseJetLink As Boolean = True) As Variant
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Variant
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT MAX(" & SafeName(Campo) & ") AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        Valor = Null
    Else
        Valor = rs!MiValor
    End If
    
    rMax = Valor
    GoTo Salir

ErrHandler:
    rMax = Null
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rMin – Mínimo remoto
' -----------------------------------------
Public Function rMin(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "", _
                     Optional ByVal DbPath As String = "", _
                     Optional ByVal UseJetLink As Boolean = True) As Variant
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Variant
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT MIN(" & SafeName(Campo) & ") AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        Valor = Null
    Else
        Valor = rs!MiValor
    End If
    
    rMin = Valor
    GoTo Salir

ErrHandler:
    rMin = Null
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rSCounter – Primer hueco disponible
' -----------------------------------------
Public Function rSCounter(ByVal Campo As String, _
                          ByVal Tabla As String, _
                          Optional ByVal MiWhere As String = "", _
                          Optional ByVal DbPath As String = "", _
                          Optional ByVal UseJetLink As Boolean = True) As Long
                          
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim Valor As Long
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    ' 1. ¿Existe el 1?
    sql = "SELECT TOP 1 " & SafeName(Campo) & " FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla) & " WHERE " & SafeName(Campo) & " = 1"
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " AND " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        rSCounter = 1
        GoTo Salir
    End If
    
    rs.Close
    Set rs = Nothing
    
    ' 2. Buscar hueco
    sql = "SELECT MIN(" & SafeName(Campo) & " + 1) AS Hueco FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla) & " WHERE NOT (" & SafeName(Campo) & " + 1) IN ("
    
    sql = sql & "SELECT " & SafeName(Campo) & " FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & ")"
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " AND " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    Valor = Nz(rs!Hueco, 1)
    rSCounter = Valor
    GoTo Salir

ErrHandler:
    rSCounter = 1
    Resume Salir

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rFirst – Primer valor remoto
' -----------------------------------------
Public Function rFirst(ByVal Campo As String, _
                       ByVal Tabla As String, _
                       Optional ByVal MiWhere As String = "", _
                       Optional ByVal DbPath As String = "", _
                       Optional ByVal UseJetLink As Boolean = True) As Variant
                       
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS Dato FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " ORDER BY " & SafeName(Campo) & " ASC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        rFirst = Null
    Else
        rFirst = rs!Dato
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
    Exit Function
    
ErrHandler:
    rFirst = Null
    Resume Salir
End Function


' -----------------------------------------
' rLast – Último valor remoto
' -----------------------------------------
Public Function rLast(ByVal Campo As String, _
                      ByVal Tabla As String, _
                      Optional ByVal MiWhere As String = "", _
                      Optional ByVal DbPath As String = "", _
                      Optional ByVal UseJetLink As Boolean = True) As Variant
                      
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & " AS Dato FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " ORDER BY " & SafeName(Campo) & " DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        rLast = Null
    Else
        rLast = rs!Dato
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
    Exit Function
    
ErrHandler:
    rLast = Null
    Resume Salir
End Function


' -----------------------------------------
' rAvg – Media remota
' -----------------------------------------
Public Function rAvg(ByVal Campo As String, _
                     ByVal Tabla As String, _
                     Optional ByVal MiWhere As String = "", _
                     Optional ByVal DbPath As String = "", _
                     Optional ByVal UseJetLink As Boolean = True) As Variant
                     
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT AVG(" & SafeName(Campo) & ") AS MiValor FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        rAvg = Null
    Else
        rAvg = Nz(rs!MiValor, Null)
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
    Exit Function
    
ErrHandler:
    rAvg = Null
    Resume Salir
End Function


' -----------------------------------------
' rMedian – Mediana remota
' -----------------------------------------
Public Function rMedian(ByVal Campo As String, _
                        ByVal Tabla As String, _
                        Optional ByVal MiWhere As String = "", _
                        Optional ByVal DbPath As String = "", _
                        Optional ByVal UseJetLink As Boolean = True) As Variant
                        
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    Dim n As Long, pos As Long
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    ' 1. Contar registros
    sql = "SELECT COUNT(" & SafeName(Campo) & ") AS N FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    If rs.EOF Then GoTo NoData
    n = Nz(rs!n, 0)
    rs.Close
    
    If n = 0 Then GoTo NoData
    
    ' 2. Obtener valores ordenados
    sql = "SELECT " & SafeName(Campo) & " AS V FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " ORDER BY " & SafeName(Campo) & " ASC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If (n Mod 2) = 1 Then
        ' Impar
        pos = (n + 1) \ 2
        rs.Move pos - 1
        rMedian = rs!v
    Else
        ' Par
        Dim v1 As Variant, V2 As Variant
        Dim Pos1 As Long, Pos2 As Long
        
        Pos1 = n \ 2
        Pos2 = Pos1 + 1
        
        rs.Move Pos1 - 1
        v1 = rs!v
        rs.MoveNext
        V2 = rs!v
        
        rMedian = (v1 + V2) / 2
    End If
    
    GoTo Salir

NoData:
    rMedian = Null
    GoTo Salir

ErrHandler:
    rMedian = Null

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


' -----------------------------------------
' rMode – Moda remota
' -----------------------------------------
Public Function rMode(ByVal Campo As String, _
                      ByVal Tabla As String, _
                      Optional ByVal MiWhere As String = "", _
                      Optional ByVal DbPath As String = "", _
                      Optional ByVal UseJetLink As Boolean = True) As Variant
                      
    Dim db As dao.Database
    Dim rs As dao.Recordset
    Dim sql As String
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    sql = "SELECT TOP 1 " & SafeName(Campo) & ", COUNT(*) AS Frecuencia FROM "
    
    If UseJetLink And DbPath <> "" Then
        sql = sql & SafeName(DbPath) & "."
    End If
    
    sql = sql & SafeName(Tabla)
    
    If Trim(MiWhere) <> "" Then
        sql = sql & " WHERE " & MiWhere
    End If
    
    sql = sql & " GROUP BY " & SafeName(Campo) & _
                " ORDER BY COUNT(*) DESC"
    
    Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
    
    If rs.EOF Then
        rMode = Null
    Else
        rMode = rs.Fields(0).Value
    End If
    
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
    Exit Function
    
ErrHandler:
    rMode = Null
    Resume Salir
End Function


