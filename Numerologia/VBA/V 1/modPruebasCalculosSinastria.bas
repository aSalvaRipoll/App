' ============================================================================
' PRUEBA DE SINASTRÍA
' ============================================================================

Public Sub PruebaSinastria()
    Dim objSinastria As clsCalculoSinastria
    
    Set objSinastria = New clsCalculoSinastria
    
    ' Configurar datos
    objSinastria.Nombre1 = "JUAN CARLOS MARTINEZ"
    objSinastria.Fecha1 = #3/15/1980#
    objSinastria.Nombre2 = "MARIA CARMEN RODRIGUEZ"
    objSinastria.Fecha2 = #7/22/1985#
    
    ' Probar los 3 tipos
    Debug.Print "=========================================="
    Debug.Print "PRUEBA DE SINASTRÍA - TIPO GENERAL"
    Debug.Print "=========================================="
    objSinastria.TipoSinastriaActual = TipoSinastria.General
    Debug.Print objSinastria.ObtenerResumenNumeros()
    Debug.Print ""
    Debug.Print objSinastria.ObtenerTodasLasRutas()
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "PRUEBA DE SINASTRÍA - TIPO ROMÁNTICA"
    Debug.Print "=========================================="
    objSinastria.TipoSinastriaActual = TipoSinastria.Romantica
    Debug.Print objSinastria.ObtenerTodasLasRutas()
    
    Debug.Print ""
    Debug.Print "=========================================="
    Debug.Print "PRUEBA DE SINASTRÍA - TIPO LABORAL"
    Debug.Print "=========================================="
    objSinastria.TipoSinastriaActual = TipoSinastria.Laboral
    Debug.Print objSinastria.ObtenerTodasLasRutas()
    
    Set objSinastria = Nothing
End Sub