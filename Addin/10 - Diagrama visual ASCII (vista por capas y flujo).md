                           +------------------------+
                           |      Access / UI       |
                           |  - Cinta (Ribbon XML)  |
                           |  - FormInspector       |
                           +-----------+------------+
                                       |
                                       v
                         +-------------+--------------+
                         |      Capa de orquestación  |
                         |        modInspectorMain    |
                         |                            |
                         |  - Inspector_Analizar      |
                         |  - Inspector_Reparar       |
                         |  - Inspector_Exportar      |
                         |  - Inspector_Resumen       |
                         |  - Inspector_Reset         |
                         |  - Inspector_Log           |
                         +-------------+--------------+
                                       |
                     +-----------------+------------------+
                     |                                    |
                     v                                    v
        +------------+-------------+        +-------------+------------+
        |       Motor de análisis  |        |      Exportación         |
        |     clsCatalogoInspector |        |   módulos ExportarXXX    |
        |                          |        |                          |
        | - AnalizarProyecto       |        | - ExportarResultadosTXT  |
        | - AnalizarModulo/Clase   |        | - ExportarResultadosXLSX |
        | - Construir catálogo     |        | - ExportarTodoHTML       |
        +------------+-------------+        +-------------+------------+
                     |
                     v
        +------------+-------------+
        |   Resultados del análisis|
        |    clsResultadosInspector|
        |                          |
        | - Colecciones internas   |
        | - RepararProyecto        |
        | - GenerarResumen         |
        +------------+-------------+
                     ^
                     |
         +-----------+-----------+
         |    Estados y mensajes |
         |                       |
         |  modInspectorEnums    |
         |  - EstadoAnalisis     |
         |  - EstadoReparacion   |
         |  - EstadoExportacion  |
         |  - FormatoExportacion |
         |  - EstiloHtml         |
         |                       |
         |  modInspectorMensajes |
         |  - MensajeAnalisis    |
         |  - MensajeReparacion  |
         |  - MensajeExportacion |
         +-----------+-----------+
                     ^
                     |
         +-----------+-----------+
         |       Infraestructura |
         |                       |
         |  - Inspector_Log      |
         |  - (futuro: logs a    |
         |     archivo / tabla)  |
         +-----------------------+
