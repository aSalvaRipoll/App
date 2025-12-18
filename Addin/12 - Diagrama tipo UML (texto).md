+--------------------------------------------------------------+
|                        modInspectorMain                      |
+--------------------------------------------------------------+
| + Inspector_Analizar() As EstadoAnalisis                     |
| + Inspector_Reparar() As EstadoReparacion                    |
| + Inspector_Exportar(formato As FormatoExportacion,          |
|                      ruta As String,                         |
|                      Optional estilo As EstiloHtml)          |
|                    As EstadoExportacion                      |
| + Inspector_Resumen() As String                              |
| + Inspector_Reset(Optional reiniciarMotor As Boolean)        |
| + Inspector_Log(mensaje As String)                           |
|                                                              |
| - gCatalogoInspector As clsCatalogoInspector                 |
| - gResultadosInspector As clsResultadosInspector             |
| - gUltimoFormato As FormatoExportacion                       |
| - gUltimaRuta As String                                      |
| - gUltimoEstiloHtml As EstiloHtml                            |
+--------------------------------------------------------------+
                |                           ^
                | uses                      | uses
                v                           |
+-----------------------------+     +---------------------------+
|      clsCatalogoInspector   |     |   clsResultadosInspector  |
+-----------------------------+     +---------------------------+
| + AnalizarProyecto() As     |     | + RepararProyecto()       |
|   clsResultadosInspector    |     | + GenerarResumen() As     |
| + (AnalizarModulo, etc.)    |     |   String                  |
+-----------------------------+     | + (colecciones internas)  |
                                    +---------------------------+


+-----------------------------+      +------------------------------+
|      modInspectorEnums      |      |     modInspectorMensajes     |
+-----------------------------+      +------------------------------+
| Enum EstadoAnalisis         |      | + MensajeAnalisis(           |
|   - AnalisisNoEjecutado     |      |     estado As EstadoAnalisis)|
|   - AnalisisEjecutado       |      |     As String                |
|   - AnalisisConErrores      |      | + MensajeReparacion(         |
| Enum EstadoReparacion       |      |     estado As EstadoReparacion|
|   - ReparacionNoEjecutada   |      |     ) As String              |
|   - ReparacionEjecutada     |      | + MensajeExportacion(        |
|   - ReparacionConErrores    |      |     estado As EstadoExportacion|
| Enum EstadoExportacion      |      |     ) As String              |
|   - ExportacionNoEjecutada  |      |                              |
|   - ExportacionEjecutada    |      | - mMensajesAnalisis          |
|   - ExportacionConErrores   |      | - mMensajesReparacion        |
| Enum FormatoExportacion     |      | - mMensajesExportacion       |
| Enum EstiloHtml             |      +------------------------------+
+-----------------------------+                ^
                ^                              |
                | uses                         | uses
                +------------------------------+


+----------------------------------------+
|             FormInspector              |
+----------------------------------------+
| + cmdAnalizar_Click()                  |
| + cmdReparar_Click()                   |
| + cmdExportar_Click()                  |
| + (otros eventos UI)                   |
+----------------------------------------+
            |           ^
            | calls     | gets messages
            v           |
   modInspectorMain     +--> modInspectorMensajes


+----------------------------------------+
|          modRibbonInspector            |
+----------------------------------------+
| + Ribbon_OnLoad(ribbon As IRibbonUI)   |
| + Ribbon_Analizar(control As ...)      |
| + Ribbon_Reparar(control As ...)       |
| + Ribbon_Exportar(control As ...)      |
| + Ribbon_LimpiarResultados(control)    |
| + Ribbon_ReiniciarMotor(control)       |
| + Ribbon_Resumen(control)              |
| + Ribbon_Invalidate()                  |
+----------------------------------------+
            |             ^
            | calls       | gets messages
            v             |
   modInspectorMain       +--> modInspectorMensajes



- Relaciones clave en formato lista:

### FormInspector

- depende de modInspectorMain (para ejecutar acciones)
- depende de modInspectorMensajes (para mostrar mensajes)

### modRibbonInspector

- depende de modInspectorMain
- depende de modInspectorMensajes

### modInspectorMain

- depende de clsCatalogoInspector
- depende de clsResultadosInspector
- depende de modInspectorEnums
- no depende de UI directa (ni forms ni ribbon)

### clsCatalogoInspector

- no depende de UI
- no depende de ribbon
- devuelve siempre clsResultadosInspector

### clsResultadosInspector

- no depende de UI
- no escribe logs por sí mismo

### modInspectorMensajes

- depende de modInspectorEnums
- no depende de UI
- sirve tanto para forms como para ribbon
