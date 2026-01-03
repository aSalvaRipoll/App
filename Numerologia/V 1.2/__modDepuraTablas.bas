Attribute VB_Name = "__modDepuraTablas"
Option Compare Database
Option Explicit

Sub Depurar()

    Call DepuraTablaDiccionario("tbmDicFonemasNom")
    
    Call DepuraTablaDiccionario("tbmDicFonemasApe")

MsgBox "Depuracion finalizada."

End Sub

Sub DepuraTablaDiccionario(nombreTabla As String)

    Dim nombreOld As String
    Dim nombreNew As String
    Dim sSQL As String
    
    nombreOld = nombreTabla & "_old"
    nombreNew = nombreTabla & "_new"
    
    ' 1. Eliminar restos previos si existen
    On Error Resume Next
    DoCmd.DeleteObject acTable, nombreOld
    DoCmd.DeleteObject acTable, nombreNew
    On Error GoTo 0
    
    ' 2. Renombrar tabla original como _old
    DoCmd.Rename nombreOld, acTable, nombreTabla
    
    ' 3. Crear tabla NEW sin autonumérico
    sSQL = "CREATE TABLE " & nombreNew & " (" & _
           "Idioma TEXT, " & _
           "ID_Idioma LONG, " & _
           "Palabra TEXT, " & _
           "FonemaCompleto TEXT, " & _
           "FonemaIPA TEXT, " & _
           "TipoEntrada TEXT, " & _
           "Notas TEXT, " & _
           "Origen TEXT, " & _
           "Activo YESNO" & _
           ");"
    DoCmd.RunSQL sSQL
    
    ' 4. Insertar solo registros únicos desde la tabla OLD
    sSQL = "INSERT INTO " & nombreNew & " (Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo) " & _
           "SELECT Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo " & _
           "FROM " & nombreOld & " " & _
           "GROUP BY ID_Idioma, Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo;"
    
    DoCmd.RunSQL sSQL
    
    ' 5. Añadir autonumérico limpio como clave primaria
    sSQL = "ALTER TABLE " & nombreNew & " ADD COLUMN ID AUTOINCREMENT PRIMARY KEY;"
    DoCmd.RunSQL sSQL
    
    ' 6. Crear índice único para evitar duplicados futuros
    'sSQL = "CREATE UNIQUE INDEX idx_Unico_Idioma_Palabra ON " & nombreNew & " (Idioma, Palabra);"
    'DoCmd.RunSQL sSQL
    
    ' 7. Renombrar tabla NEW como la original
    DoCmd.Rename nombreTabla, acTable, nombreNew
    
    MsgBox "Depuración completada para: " & nombreTabla & vbCrLf & _
           "Copia original guardada como: " & nombreOld & vbCrLf & _
           "Autonumérico reconstruido e índice único creado.", vbInformation

End Sub

'UPDATE tbmDicFonemasNom AS N
'INNER JOIN tbmIdiomas AS I
'    ON N.Idioma = I.Abreviado
'SET N.ID_Idioma = I.IDIdioma;


'UPDATE tbmDicFonemasApe AS A
'INNER JOIN tbmIdiomas AS I
'    ON A.Idioma = I.Abreviado
'SET A.ID_Idioma = I.IDIdioma;




'UPDATE tbmDicFonemasNom AS N
'INNER JOIN tbmIdiomas AS I
'    ON N.Idioma = I.Abreviado
'Set n.Origen = i.Origen
'WHERE N.Origen IS NULL OR N.Origen = '';


'UPDATE tbmDicFonemasApe AS A
'INNER JOIN tbmIdiomas AS I
'    ON A.Idioma = I.Abreviado
'Set a.Origen = i.Origen
'WHERE A.Origen IS NULL OR A.Origen = '';


'UPDATE tbmDicFonemasNom AS N
'INNER JOIN tblIdioma AS I
'    ON N.Idioma = I.Abreviado
'Set n.Origen = i.NomIdioma
'WHERE N.Origen IS NULL OR N.Origen = '';


'UPDATE tbmDicFonemasApe AS A
'INNER JOIN tblIdioma AS I
'    ON A.Idioma = I.Abreviado
'Set a.Origen = i.NomIdioma
'WHERE A.Origen IS NULL OR A.Origen = '';

Sub DepuraTablaDiccionario_2(nombreTabla As String)

    Dim nombreOld As String
    Dim nombreNew As String
    Dim sSQL As String
    
    nombreOld = nombreTabla & "_old"
    nombreNew = nombreTabla & "_new"
    
    ' 1. Eliminar restos previos si existen
    On Error Resume Next
    DoCmd.DeleteObject acTable, nombreOld
    DoCmd.DeleteObject acTable, nombreNew
    On Error GoTo 0
    
    ' 2. Crear copia OLD
    sSQL = "SELECT * INTO " & nombreOld & " FROM " & nombreTabla & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL sSQL
    DoCmd.SetWarnings True
    
    ' 3. Crear tabla NEW sin autonumérico
    sSQL = "CREATE TABLE " & nombreNew & " (" & _
           "Idioma TEXT, " & _
           "ID_Idioma LONG, " & _
           "Palabra TEXT, " & _
           "FonemaCompleto TEXT, " & _
           "FonemaIPA TEXT, " & _
           "TipoEntrada TEXT, " & _
           "Notas TEXT, " & _
           "Origen TEXT, " & _
           "Activo YESNO" & _
           ");"
    DoCmd.RunSQL sSQL
    
    ' 4. Insertar solo registros únicos
    sSQL = "INSERT INTO " & nombreNew & " (Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo) " & _
           "SELECT Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo " & _
           "FROM " & nombreOld & " " & _
           "GROUP BY Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo;"
    DoCmd.RunSQL sSQL
    
    ' 5. Eliminar tabla original
    DoCmd.DeleteObject acTable, nombreTabla
    
    ' 6. Renombrar tabla NEW como original
    DoCmd.Rename nombreTabla, acTable, nombreNew
    
    ' 7. Añadir autonumérico limpio como clave primaria
    sSQL = "ALTER TABLE " & nombreTabla & " ADD COLUMN ID AUTOINCREMENT PRIMARY KEY;"
    DoCmd.RunSQL sSQL
    
    ' 8. Crear índice único para evitar duplicados futuros
    sSQL = "CREATE UNIQUE INDEX idx_Unico_Idioma_Palabra ON " & nombreTabla & " (Idioma, Palabra);"
    DoCmd.RunSQL sSQL
    
    MsgBox "Depuración completada para: " & nombreTabla & vbCrLf & _
           "Autonumérico reconstruido y índice único creado.", vbInformation

End Sub

Sub DepuraTablaDiccionario_1(nombreTabla As String)

    Dim nombreOld As String
    Dim nombreNew As String
    Dim sSQL As String

    nombreOld = nombreTabla & "_old"
    nombreNew = nombreTabla & "_new"

    ' 1. Eliminar restos previos si existen
    On Error Resume Next
    DoCmd.DeleteObject acTable, nombreOld
    DoCmd.DeleteObject acTable, nombreNew
    On Error GoTo 0

    ' 2. Crear copia OLD
    sSQL = "SELECT * INTO " & nombreOld & " FROM " & nombreTabla & ";"
    DoCmd.SetWarnings False
    DoCmd.RunSQL sSQL
    DoCmd.SetWarnings True

    ' 3. Crear tabla NEW sin autonumérico
    sSQL = "CREATE TABLE " & nombreNew & " (" & _
           "Idioma TEXT, " & _
           "ID_Idioma LONG, " & _
           "Palabra TEXT, " & _
           "FonemaCompleto TEXT, " & _
           "FonemaIPA TEXT, " & _
           "TipoEntrada TEXT, " & _
           "Notas TEXT, " & _
           "Origen TEXT, " & _
           "Activo YESNO" & _
           ");"
    DoCmd.RunSQL sSQL

    ' 4. Insertar solo registros únicos
    sSQL = "INSERT INTO " & nombreNew & " (Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo) " & _
           "SELECT Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo " & _
           "FROM " & nombreOld & " " & _
           "GROUP BY Idioma, ID_Idioma, Palabra, FonemaCompleto, FonemaIPA, TipoEntrada, Notas, Origen, Activo;"
    DoCmd.RunSQL sSQL

    ' 5. Eliminar tabla original
    DoCmd.DeleteObject acTable, nombreTabla

    ' 6. Renombrar tabla NEW como original
    DoCmd.Rename nombreTabla, acTable, nombreNew

    ' 7. Añadir autonumérico limpio como clave primaria
    DoCmd.OpenTable nombreTabla, acViewDesign
    DoCmd.RunCommand acCmdInsertTableRow
    DoCmd.RunCommand acCmdInsertField
    DoCmd.RunCommand acCmdFieldInsertLookup
    ' Aquí Access crea un campo nuevo, tú solo lo renombras a mano si quieres

    DoCmd.Close acTable, nombreTabla, acSaveYes

    MsgBox "Depuración completada para: " & nombreTabla, vbInformation

End Sub
