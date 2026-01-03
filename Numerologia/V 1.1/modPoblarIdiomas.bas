Attribute VB_Name = "modPoblarIdiomas"
Option Compare Database
Option Explicit

Public Sub PoblarIdiomas()

    Dim db As DAO.Database
    Set db = CurrentDb

    ' Opcional: limpiar la tabla antes de poblarla
    db.Execute "DELETE FROM tbmIdiomas", dbFailOnError

    ' Insertar idioma global / otros (ID = 0)
    db.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) " & _
               "VALUES (0, 'other', 'Otros / Global', 'Idiomas no contemplados')", dbFailOnError

    ' Insertar idiomas principales
    db.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) " & _
               "VALUES (1, 'es', 'Castellano', 'Español estándar')", dbFailOnError

    db.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) " & _
               "VALUES (2, 'ca', 'Català', 'Catalán')", dbFailOnError

    db.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) " & _
               "VALUES (3, 'eu', 'Euskara', 'Vasco')", dbFailOnError

    db.Execute "INSERT INTO tbmIdiomas (IDIdioma, Abreviado, NomIdioma, Notas) " & _
               "VALUES (4, 'gl', 'Galego', 'Gallego')", dbFailOnError

    

    MsgBox "Tabla tbmIdiomas poblada correctamente.", vbInformation

End Sub


