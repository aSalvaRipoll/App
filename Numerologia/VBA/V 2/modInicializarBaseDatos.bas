Attribute VB_Name = "modInicializarBaseDatos"
Public Sub InicializarBaseDatos()
    ' Crear tablas
    Call CrearTodasLasTablas
    
    ' Cargar datos
    Call CargarTodosLosDatos
    
    MsgBox "Base de datos inicializada correctamente", vbInformation, "Completado"
End Sub
