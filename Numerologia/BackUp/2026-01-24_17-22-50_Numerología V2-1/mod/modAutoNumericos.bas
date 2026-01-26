Attribute VB_Name = "modAutoNumericos"

Option Compare Database
Option Explicit

' ============================================================================
'   Módulo: modAutoNumericos
'   Autor:  Alba Salvá
'   Año:    2026
'
'   Propósito:
'       Generación de autonuméricos controlados, coherentes y auditables.
'       Sustituye completamente los autonuméricos nativos de Access.
'
'       Función principal:
'           AutoNext --> genera el siguiente número según dos modos:
'               - RellenarHuecos = True  --> primer hueco disponible
'               - RellenarHuecos = False --> secuencial puro
'
'       Compatible con:
'           - Access local
'           - Bases externas Access
'           - JetLink
'           - SQL Server, Oracle, MySQL, PostgreSQL (vía ODBC)
'
'   Notas:
'       - No cierra CurrentDb nunca.
'       - Limpieza segura de objetos.
'       - SQL seguro con SafeName.
' ============================================================================


' ============================================================================
'   Funciones auxiliares
' ============================================================================

Public Function SafeName(ByVal Nombre As String) As String
    SafeName = "[" & Replace(Nombre, "]", "]]") & "]"
End Function


Private Function GetDB(ByVal DbPath As String, ByVal UseJetLink As Boolean) As DAO.Database
    If DbPath = "" Or (DbPath <> "" And UseJetLink) Then
        Set GetDB = CurrentDb
    Else
        Set GetDB = DBEngine.OpenDatabase(DbPath)
    End If
End Function



' ============================================================================
'   FUNCIÓN PRINCIPAL DE NUMERACIÓN
' ============================================================================

Public Function AutoNext(ByVal Campo As String, _
                         ByVal Tabla As String, _
                         Optional ByVal MiWhere As String = "", _
                         Optional ByVal DbPath As String = "", _
                         Optional ByVal UseJetLink As Boolean = True, _
                         Optional ByVal RellenarHuecos As Boolean = True) As Long
                         
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim Valor As Long
'    Dim N As Long
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    ' ----------------------------------------------------------
    ' 1. MODO SECUENCIAL PURO (sin rellenar huecos)
    ' ----------------------------------------------------------
    If Not RellenarHuecos Then
        
        sql = "SELECT MAX(" & SafeName(Campo) & ") AS MaxVal FROM "
        
        If UseJetLink And DbPath <> "" Then
            sql = sql & SafeName(DbPath) & "."
        End If
        
        sql = sql & SafeName(Tabla)
        
        If Trim(MiWhere) <> "" Then
            sql = sql & " WHERE " & MiWhere
        End If
        
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
        
        If rs.EOF Then
            AutoNext = 1
        Else
            AutoNext = Nz(rs!maxVal, 0) + 1
        End If
        
        GoTo Salir
    End If
    
    
    ' ----------------------------------------------------------
    ' 2. MODO RELLENAR HUECOS (primer hueco disponible)
    ' ----------------------------------------------------------
    
    ' 2.1 ¿Existe el 1?
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
        AutoNext = 1
        GoTo Salir
    End If
    
    rs.Close
    Set rs = Nothing
    
    
    ' 2.2 Buscar el primer hueco
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
    AutoNext = Valor
    GoTo Salir
    
    
' ----------------------------------------------------------
' 3. Manejo de errores
' ----------------------------------------------------------
ErrHandler:
    AutoNext = 1
    Resume Salir

    
' ----------------------------------------------------------
' 4. Limpieza
' ----------------------------------------------------------
Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function


'====

Private Function SqlLiteral(ByVal v As Variant) As String
    If IsNull(v) Then
        SqlLiteral = "Null"
    ElseIf VarType(v) = vbString Then
        SqlLiteral = "'" & Replace(v, "'", "''") & "'"
    ElseIf VarType(v) = vbDate Then
        SqlLiteral = "#" & Format$(v, "yyyy/mm/dd hh:nn:ss") & "#"
    Else
        SqlLiteral = CStr(v)
    End If
End Function

'----

Public Function AutoNextSafe(ByVal Campo As String, _
                             ByVal Tabla As String, _
                             Optional ByVal MiWhere As String = "", _
                             Optional ByVal DbPath As String = "", _
                             Optional ByVal UseJetLink As Boolean = True, _
                             Optional ByVal RellenarHuecos As Boolean = True, _
                             Optional ByVal MaxReintentos As Long = 5) As Long
                             
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim Intento As Long
    Dim n As Long
    
    On Error GoTo ErrHandler
    
    Set db = GetDB(DbPath, UseJetLink)
    
    For Intento = 1 To MaxReintentos
        
        'db.BeginTrans
        
        n = AutoNext(Campo, Tabla, MiWhere, DbPath, UseJetLink, RellenarHuecos)
        
        ' Verificar que no exista ya
        sql = "SELECT " & SafeName(Campo) & " FROM "
        If UseJetLink And DbPath <> "" Then sql = sql & SafeName(DbPath) & "."
        sql = sql & SafeName(Tabla) & " WHERE " & SafeName(Campo) & " = " & CStr(n)
        If Trim(MiWhere) <> "" Then sql = sql & " AND " & MiWhere
        
        Set rs = db.OpenRecordset(sql, dbOpenSnapshot, dbReadOnly)
        
        If rs.EOF Then
            ' Número libre --> OK
            'db.CommitTrans
            AutoNextSafe = n
            GoTo Salir
        Else
            ' Colisión --> deshacer y reintentar
            'db.Rollback
        End If
        
        rs.Close
        Set rs = Nothing
    Next Intento
    
    ' Si llega aquí, no ha podido encontrar número seguro
    AutoNextSafe = 0
    GoTo Salir

ErrHandler:
    On Error Resume Next
    'db.Rollback
    AutoNextSafe = 0

Salir:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If DbPath <> "" And Not UseJetLink Then db.Close
    Set db = Nothing
End Function

'----

Public Function AutoNextByYear(ByVal Campo As String, _
                               ByVal Tabla As String, _
                               ByVal CampoAnyo As String, _
                               Optional ByVal Anyo As Long = 0, _
                               Optional ByVal DbPath As String = "", _
                               Optional ByVal UseJetLink As Boolean = True, _
                               Optional ByVal RellenarHuecos As Boolean = True) As Long
                               
    Dim Filtro As String
    Dim ValorAnyo As Long
    
    If Anyo = 0 Then
        ValorAnyo = Year(Date)
    Else
        ValorAnyo = Anyo
    End If
    
    Filtro = SafeName(CampoAnyo) & " = " & CStr(ValorAnyo)
    
    AutoNextByYear = AutoNext(Campo, Tabla, Filtro, DbPath, UseJetLink, RellenarHuecos)
End Function

'----

Public Function AutoNextByField(ByVal Campo As String, _
                                ByVal Tabla As String, _
                                ByVal CampoClave As String, _
                                ByVal ValorClave As Variant, _
                                Optional ByVal DbPath As String = "", _
                                Optional ByVal UseJetLink As Boolean = True, _
                                Optional ByVal RellenarHuecos As Boolean = True) As Long
                                
    Dim Filtro As String
    
    Filtro = SafeName(CampoClave) & " = " & SqlLiteral(ValorClave)
    
    AutoNextByField = AutoNext(Campo, Tabla, Filtro, DbPath, UseJetLink, RellenarHuecos)
End Function

'----

Public Function AutoNextPrefix(ByVal CampoNumero As String, _
                               ByVal CampoPrefijo As String, _
                               ByVal Tabla As String, _
                               ByVal Prefijo As String, _
                               Optional ByVal ancho As Long = 6, _
                               Optional ByVal DbPath As String = "", _
                               Optional ByVal UseJetLink As Boolean = True, _
                               Optional ByVal RellenarHuecos As Boolean = True) As String
                               
    Dim Filtro As String
    Dim n As Long
    
    Filtro = SafeName(CampoPrefijo) & " = " & SqlLiteral(Prefijo)
    
    n = AutoNext(CampoNumero, Tabla, Filtro, DbPath, UseJetLink, RellenarHuecos)
    
    AutoNextPrefix = Prefijo & "-" & Format$(n, String$(ancho, "0"))
End Function

'----

Public Function AutoNextComposite(ByVal CampoNumero As String, _
                                  ByVal CampoAnyo As String, _
                                  ByVal CampoSerie As String, _
                                  ByVal Tabla As String, _
                                  ByVal Serie As String, _
                                  Optional ByVal Anyo As Long = 0, _
                                  Optional ByVal ancho As Long = 6, _
                                  Optional ByVal DbPath As String = "", _
                                  Optional ByVal UseJetLink As Boolean = True, _
                                  Optional ByVal RellenarHuecos As Boolean = True) As String
                                  
    Dim Filtro As String
    Dim ValorAnyo As Long
    Dim n As Long
    
    If Anyo = 0 Then
        ValorAnyo = Year(Date)
    Else
        ValorAnyo = Anyo
    End If
    
    Filtro = SafeName(CampoAnyo) & " = " & CStr(ValorAnyo) & _
             " AND " & SafeName(CampoSerie) & " = " & SqlLiteral(Serie)
    
    n = AutoNext(CampoNumero, Tabla, Filtro, DbPath, UseJetLink, RellenarHuecos)
    
    AutoNextComposite = CStr(ValorAnyo) & "-" & Serie & "-" & Format$(n, String$(ancho, "0"))
End Function

'----

Public Function AutoNextDaily(ByVal CampoNumero As String, _
                              ByVal CampoFecha As String, _
                              ByVal Tabla As String, _
                              Optional ByVal Fecha As Date = 0, _
                              Optional ByVal ancho As Long = 4, _
                              Optional ByVal DbPath As String = "", _
                              Optional ByVal UseJetLink As Boolean = True, _
                              Optional ByVal RellenarHuecos As Boolean = True) As String
                              
    Dim f As Date
    Dim Filtro As String
    Dim n As Long
    
    If Fecha = 0 Then
        f = Date
    Else
        f = Fecha
    End If
    
    Filtro = SafeName(CampoFecha) & " = #" & Format$(f, "yyyy/mm/dd") & "#"
    
    n = AutoNext(CampoNumero, Tabla, Filtro, DbPath, UseJetLink, RellenarHuecos)
    
    AutoNextDaily = Format$(f, "yyyy/mm/dd") & "-" & Format$(n, String$(ancho, "0"))
End Function


