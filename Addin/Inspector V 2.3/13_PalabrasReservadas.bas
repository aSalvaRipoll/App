Attribute VB_Name = "13_PalabrasReservadas"

'===============================================================
' Módulo: 13_PalabrasReservadas
' Gestión centralizada de palabras reservadas (VBA, SQL, Access)
'===============================================================

Option Compare Database
Option Explicit

'-------------------------------------------------------
'  Variables internas
'-------------------------------------------------------

Private m_AllWords As Collection
Private m_Initialized As Boolean

'-------------------------------------------------------
'  API pública
'-------------------------------------------------------

Public Function GetAllReservedWords() As Collection
    EnsureInitialized
    Set GetAllReservedWords = m_AllWords
End Function

Public Function GetReservedWordsByCategory(ByVal cat As ReservedCategory) As Collection
    EnsureInitialized

    Dim col As New Collection
    Dim item As Variant
    Dim info As ReservedWordInfo

    For Each item In m_AllWords
        info = item
        If info.Categoria = cat Then
            col.Add info
        End If
    Next item

    Set GetReservedWordsByCategory = col
End Function

Public Function FindReservedWord(ByVal nombre As String) As Collection
    EnsureInitialized

    Dim col As New Collection
    Dim item As Variant
    Dim info As ReservedWordInfo
    Dim sName As String

    sName = LCase$(Trim$(nombre))

    For Each item In m_AllWords
        info = item
        If LCase$(info.nombre) = sName Then
            col.Add info
        End If
    Next item

    Set FindReservedWord = col
End Function

'-------------------------------------------------------
'  Inicialización
'-------------------------------------------------------

Private Sub EnsureInitialized()
    If m_Initialized Then Exit Sub

    Set m_AllWords = New Collection
    LoadFromTable

    m_Initialized = True
End Sub

'-------------------------------------------------------
'  Carga desde tabla
'-------------------------------------------------------

Private Sub LoadFromTable()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim info As ReservedWordInfo

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Nombre, Tipo, Categoria FROM tblPalabrasReservadas WHERE Activa = True")

    Do While Not rs.EOF
        info.nombre = Nz(rs!nombre, "")
        info.Tipo = Nz(rs!Tipo, "")
        info.Categoria = Nz(rs!Categoria, 0)

        m_AllWords.Add info
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

'=========================================================================


Public Sub ImportarPalabrasDesdeFicheros()
    Dim fso As Object
    Dim carpeta As Object
    Dim archivo As Object
    Dim ruta As String
    Dim linea As String
    Dim partes() As String
    Dim cat As ReservedCategory
    Dim db As DAO.Database
    Dim sql As String
    Dim ff As Integer

    ruta = SeleccionarCarpeta()
    If ruta = "" Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set carpeta = fso.GetFolder(ruta)
    Set db = CurrentDb

    For Each archivo In carpeta.Files
        If LCase$(fso.GetExtensionName(archivo.Name)) = "txt" Then

            cat = CategoriaFromFileName(fso.GetBaseName(archivo.Name))

            ff = FreeFile
            Open archivo.Path For Input As #ff

            Do While Not EOF(ff)
                Line Input #ff, linea
                linea = Trim$(linea)
                If linea <> "" Then
                    partes = Split(linea, ";")
                    If UBound(partes) = 1 Then
                        sql = "INSERT INTO tblPalabrasReservadas (Nombre, Tipo, Categoria, Activa) " & _
                              "VALUES (" & _
                              "'" & Replace(partes(0), "'", "''") & "', " & _
                              "'" & Replace(partes(1), "'", "''") & "', " & _
                              cat & ", True)"
                        db.Execute sql, dbFailOnError
                    End If
                End If
            Loop

            Close #ff
        End If
    Next archivo

    MsgBox "Importación completada.", vbInformation
End Sub

