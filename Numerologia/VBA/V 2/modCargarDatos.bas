Attribute VB_Name = "modCargarDatos"
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Módulo: modCargarDatos
' Descripción: Carga los datos iniciales en las tablas
' Autor: Alba Salvá
' Fecha: 2025
' ============================================================================

Public Sub CargarTodosLosDatos()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== INICIANDO CARGA DE DATOS ==="
    Debug.Print ""
    
    ' Cargar datos de catálogos
    Call CargarTiposCalculo
    Call CargarTiposSinastria
    Call CargarConfiguracion
    
    Debug.Print ""
    Debug.Print "=== DATOS CARGADOS EXITOSAMENTE ==="
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "ERROR al cargar datos: " & err.Description
    MsgBox "Error al cargar datos: " & err.Description, vbCritical, "Error"
End Sub

' ============================================================================
' DATOS: tblTiposCalculo
' ============================================================================

Private Sub CargarTiposCalculo()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("tblTiposCalculo", dbOpenDynaset)
    
    ' Limpiar tabla
    db.Execute "DELETE FROM tblTiposCalculo", dbFailOnError
    
    ' Números básicos
    rst.AddNew
    rst!TipoCalculoID = 1
    rst!nombre = "Camino de Vida"
    rst!Descripcion = "Número principal que define el propósito de vida"
    rst!CarpetaInterpretaciones = "CaminoVida"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 2
    rst!nombre = "Destino"
    rst!Descripcion = "Misión y talentos naturales"
    rst!CarpetaInterpretaciones = "Destino"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 3
    rst!nombre = "Alma"
    rst!Descripcion = "Deseos internos y motivaciones"
    rst!CarpetaInterpretaciones = "Alma"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 4
    rst!nombre = "Personalidad"
    rst!Descripcion = "Cómo te perciben los demás"
    rst!CarpetaInterpretaciones = "Personalidad"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 5
    rst!nombre = "Madurez"
    rst!Descripcion = "Desarrollo en la segunda mitad de la vida"
    rst!CarpetaInterpretaciones = "Madurez"
    rst.Update
    
    ' Números temporales
    rst.AddNew
    rst!TipoCalculoID = 6
    rst!nombre = "Año Personal"
    rst!Descripcion = "Energía del año actual"
    rst!CarpetaInterpretaciones = "AnoPersonal"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 7
    rst!nombre = "Mes Personal"
    rst!Descripcion = "Energía del mes actual"
    rst!CarpetaInterpretaciones = "MesPersonal"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 8
    rst!nombre = "Día Personal"
    rst!Descripcion = "Energía del día actual"
    rst!CarpetaInterpretaciones = "DiaPersonal"
    rst.Update
    
    ' Ciclos
    rst.AddNew
    rst!TipoCalculoID = 9
    rst!nombre = "Ciclo 1"
    rst!Descripcion = "Primer ciclo de vida (formativo)"
    rst!CarpetaInterpretaciones = "Ciclo1"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 10
    rst!nombre = "Ciclo 2"
    rst!Descripcion = "Segundo ciclo de vida (productivo)"
    rst!CarpetaInterpretaciones = "Ciclo2"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 11
    rst!nombre = "Ciclo 3"
    rst!Descripcion = "Tercer ciclo de vida (cosecha)"
    rst!CarpetaInterpretaciones = "Ciclo3"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 12
    rst!nombre = "Ciclo 4"
    rst!Descripcion = "Cuarto ciclo de vida (espiritual)"
    rst!CarpetaInterpretaciones = "Ciclo4"
    rst.Update
    
    ' Pináculos
    rst.AddNew
    rst!TipoCalculoID = 13
    rst!nombre = "Pináculo 1"
    rst!Descripcion = "Primera meta de vida"
    rst!CarpetaInterpretaciones = "Pinaculo1"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 14
    rst!nombre = "Pináculo 2"
    rst!Descripcion = "Segunda meta de vida"
    rst!CarpetaInterpretaciones = "Pinaculo2"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 15
    rst!nombre = "Pináculo 3"
    rst!Descripcion = "Tercera meta de vida"
    rst!CarpetaInterpretaciones = "Pinaculo3"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 16
    rst!nombre = "Pináculo 4"
    rst!Descripcion = "Cuarta meta de vida"
    rst!CarpetaInterpretaciones = "Pinaculo4"
    rst.Update
    
    ' Desafíos
    rst.AddNew
    rst!TipoCalculoID = 17
    rst!nombre = "Desafío 1"
    rst!Descripcion = "Primer desafío de vida"
    rst!CarpetaInterpretaciones = "Desafio1"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 18
    rst!nombre = "Desafío 2"
    rst!Descripcion = "Segundo desafío de vida"
    rst!CarpetaInterpretaciones = "Desafio2"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 19
    rst!nombre = "Desafío 3"
    rst!Descripcion = "Tercer desafío de vida (mayor)"
    rst!CarpetaInterpretaciones = "Desafio3"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 20
    rst!nombre = "Desafío 4"
    rst!Descripcion = "Cuarto desafío de vida"
    rst!CarpetaInterpretaciones = "Desafio4"
    rst.Update
    
    ' Números especiales
    rst.AddNew
    rst!TipoCalculoID = 21
    rst!nombre = "Número de Expresión"
    rst!Descripcion = "Manera de expresarse en el mundo"
    rst!CarpetaInterpretaciones = "Expresion"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 22
    rst!nombre = "Número de Poder"
    rst!Descripcion = "Primera letra + Primera vocal"
    rst!CarpetaInterpretaciones = "Poder"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 23
    rst!nombre = "Número Faltante"
    rst!Descripcion = "Números ausentes en el nombre"
    rst!CarpetaInterpretaciones = "Ausente"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 24
    rst!nombre = "Número Dominante"
    rst!Descripcion = "Número más repetido en el nombre"
    rst!CarpetaInterpretaciones = "Dominante"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 25
    rst!nombre = "Primera Letra"
    rst!Descripcion = "Piedra angular del nombre"
    rst!CarpetaInterpretaciones = "PrimeraLetra"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 26
    rst!nombre = "Primera Vocal"
    rst!Descripcion = "Primera vocal del nombre"
    rst!CarpetaInterpretaciones = "PrimeraVocal"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 27
    rst!nombre = "Primera Consonante"
    rst!Descripcion = "Primera consonante del nombre"
    rst!CarpetaInterpretaciones = "PrimeraConsonante"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 28
    rst!nombre = "Respuesta Subconsciente"
    rst!Descripcion = "Cantidad de números presentes"
    rst!CarpetaInterpretaciones = "RespuestaSubconsciente"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 29
    rst!nombre = "Plano de Expresión"
    rst!Descripcion = "Distribución mental/físico/emocional/intuitivo"
    rst!CarpetaInterpretaciones = "PlanoExpresion"
    rst.Update
    
    rst.AddNew
    rst!TipoCalculoID = 30
    rst!nombre = "Edad Personal"
    rst!Descripcion = "Número de la edad actual"
    rst!CarpetaInterpretaciones = "EdadPersonal"
    rst.Update
    
    rst.Close
    
    Debug.Print "? Tipos de cálculo cargados (30 tipos)"
    
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not rst Is Nothing Then rst.Close
    Debug.Print "? Error al cargar tipos de cálculo: " & err.Description
End Sub

' ============================================================================
' DATOS: tblTiposSinastria
' ============================================================================

Private Sub CargarTiposSinastria()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("tblTiposSinastria", dbOpenDynaset)
    
    ' Limpiar tabla
    db.Execute "DELETE FROM tblTiposSinastria", dbFailOnError
    
    rst.AddNew
    rst!TipoSinastriaID = 1
    rst!nombre = "General"
    rst!Descripcion = "Compatibilidad general (amistad, familia)"
    rst!CarpetaInterpretaciones = "General"
    rst.Update
    
    rst.AddNew
    rst!TipoSinastriaID = 2
    rst!nombre = "Romántica"
    rst!Descripcion = "Compatibilidad romántica y de pareja"
    rst!CarpetaInterpretaciones = "Romantica"
    rst.Update
    
    rst.AddNew
    rst!TipoSinastriaID = 3
    rst!nombre = "Laboral"
    rst!Descripcion = "Compatibilidad laboral y de negocios"
    rst!CarpetaInterpretaciones = "Laboral"
    rst.Update
    
    rst.Close
    
    Debug.Print "? Tipos de sinastría cargados (3 tipos)"
    
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not rst Is Nothing Then rst.Close
    Debug.Print "? Error al cargar tipos de sinastría: " & err.Description
End Sub

' ============================================================================
' DATOS: tblConfiguracion
' ============================================================================

Private Sub CargarConfiguracion()
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
    On Error GoTo ErrorHandler
    
    Set db = CurrentDb
    Set rst = db.OpenRecordset("tblConfiguracion", dbOpenDynaset)
    
    ' Limpiar tabla
    db.Execute "DELETE FROM tblConfiguracion", dbFailOnError
    
    ' Rutas
    rst.AddNew
    rst!clave = "RutaInterpretaciones"
    rst!valor = CurrentProject.Path & "\Interpretaciones\"
    rst!Descripcion = "Ruta base de archivos de interpretaciones"
    rst.Update
    
    rst.AddNew
    rst!clave = "RutaSinastrias"
    rst!valor = CurrentProject.Path & "\Sinastrias\"
    rst!Descripcion = "Ruta base de archivos de sinastría"
    rst.Update
    
    ' Configuración de cálculos
    rst.AddNew
    rst!clave = "PermitirNumerosMaestros"
    rst!valor = "Si"
    rst!Descripcion = "Permitir números maestros (11, 22, 33, 44)"
    rst.Update
    
    rst.AddNew
    rst!clave = "PermitirNumerosKarmicos"
    rst!valor = "Si"
    rst!Descripcion = "Permitir números kármicos (13, 14, 16, 19)"
    rst.Update
    
    rst.AddNew
    rst!clave = "SistemaNumerico"
    rst!valor = "Pitagorico"
    rst!Descripcion = "Sistema numerológico a utilizar"
    rst.Update
    
    ' Interfaz
    rst.AddNew
    rst!clave = "MostrarInterpretacionesCompletas"
    rst!valor = "Si"
    rst!Descripcion = "Mostrar interpretaciones completas o resumidas"
    rst.Update
    
    rst.AddNew
    rst!clave = "IdiomaInterfaz"
    rst!valor = "Español"
    rst!Descripcion = "Idioma de la interfaz"
    rst.Update
    
    ' Versión
    rst.AddNew
    rst!clave = "VersionAplicacion"
    rst!valor = "1.0.0"
    rst!Descripcion = "Versión actual de la aplicación"
    rst.Update
    
    rst.AddNew
    rst!clave = "FechaCreacionBD"
    rst!valor = Format(Date, "dd/mm/yyyy")
    rst!Descripcion = "Fecha de creación de la base de datos"
    rst.Update
    
    rst.Close
    
    Debug.Print "? Configuración cargada (9 parámetros)"
    
    Set rst = Nothing
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    If Not rst Is Nothing Then rst.Close
    Debug.Print "? Error al cargar configuración: " & err.Description
End Sub
