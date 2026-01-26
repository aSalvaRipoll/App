Attribute VB_Name = "modSalidaHTML"

Option Compare Database
Option Explicit

' Asumo que ya tienes:
' - LeerArchivoUTF8
' - EscribirUTF8
' - ConvertirMarkdownAHTML
' - ReemplazarMarcadores
' - RellenarMarcadoresCiclos
' - RellenarMarcadoresPinaculosEscollos
' - RellenarMarcadoresTransitos
' - RellenarMarcadoresProgresiones
' - RellenarMarcadoresCasas
' - ExportarWordAPDF

Public Sub GenerarInformeNumerologico(r As clsResultado, _
                                      ByVal rutaPlantilla As String, _
                                      ByVal rutaCarpetaMD As String, _
                                      ByVal rutaHTML As String, _
                                      ByVal rutaPDF As String)

    Dim html As String
    Dim contenido As String
    Dim frag As String

    ' 1. Cargar plantilla HTML base
    html = LeerArchivoUTF8(rutaPlantilla)

    ' 2. Rellenar cabecera
    html = ReemplazarMarcadores(html, "NOMBRE", r.Nombre)
    html = ReemplazarMarcadores(html, "FECHA_NAC", r.FechaNacimiento)
    html = ReemplazarMarcadores(html, "EDAD", r.Edad)
    html = ReemplazarMarcadores(html, "SISTEMA_FONETICO", r.SistemaFonetico)
    html = ReemplazarMarcadores(html, "SISTEMA_CALCULO", r.SistemaCalculo)
    html = ReemplazarMarcadores(html, "NUM_CICLOS", r.NumCiclos)
    html = ReemplazarMarcadores(html, "METODO_CICLOS", r.MetodoCiclos)
    html = ReemplazarMarcadores(html, "SISTEMA_TAROT", r.SistemaTarot)
    html = ReemplazarMarcadores(html, "FECHA_CALCULO", r.FechaCalculo)
    html = ReemplazarMarcadores(html, "VERSION_MOTOR", r.VersionMotor)

    contenido = ""

    ' ============================
    ' 3. NÚMEROS PRINCIPALES
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\numeros_principales.md")
    frag = RellenarMarcadoresNumerosPrincipales(frag, r)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 4. CICLOS
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\ciclos.md")
    frag = RellenarMarcadoresCiclos(frag, _
                                    r.Ciclo1, r.Ciclo2, r.Ciclo3, _
                                    r.CicloRango1, r.CicloRango2, r.CicloRango3, _
                                    r.CicloTextoCorto1, r.CicloTextoCorto2, r.CicloTextoCorto3, _
                                    r.CicloTextoLargo1, r.CicloTextoLargo2, r.CicloTextoLargo3)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 5. PINÁCULOS Y ESCOLLOS
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\pinaculos.md")
    frag = RellenarMarcadoresPinaculosEscollos(frag, _
                                               r.Pinaculo1, r.Pinaculo2, r.Pinaculo3, r.Pinaculo4, _
                                               r.Escollos1, r.Escollos2, r.Escollos3, r.Escollos4, _
                                               r.TextoPinaculo1, r.TextoPinaculo2, r.TextoPinaculo3, r.TextoPinaculo4, _
                                               r.TextoEscollos1, r.TextoEscollos2, r.TextoEscollos3, r.TextoEscollos4)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 6. TRÁNSITOS ACTUALES
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\transitos.md")
    frag = RellenarMarcadoresTransitos(frag, _
                                       r.TransitoFisicoActual, _
                                       r.TransitoMentalActual, _
                                       r.TransitoEspiritualActual, _
                                       r.EsenciaAnualActual)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 7. PROGRESIONES DEL NOMBRE
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\progresiones.md")
    frag = RellenarMarcadoresProgresiones(frag, _
                                          r.ED, r.TF, r.TM, r.TE, r.EDD, r.EDS, r.AP)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 8. CASAS NUMEROLÓGICAS
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\casas.md")
    frag = RellenarMarcadoresCasas(frag, _
                                   r.CasasValor, r.CasasPorcentaje, r.CasasMedia, _
                                   r.CasasTexto, _
                                   r.TextoValores, r.TextoPorcentajes, r.TextoMedias)
    contenido = contenido & frag & vbCrLf

    ' ============================
    ' 9. INTERPRETACIÓN DEL NOMBRE
    ' ============================
    frag = SeccionDesdeMD(rutaCarpetaMD & "\interpretacion_nombre.md")
    frag = Replace(frag, "{{INTERPRETACION_NOMBRE}}", r.InterpretacionNombre)
    contenido = contenido & frag & vbCrLf

    ' 10. Insertar todo el contenido en la plantilla
    html = ReemplazarMarcadores(html, "CONTENIDO", contenido)

    ' 11. Guardar HTML final en UTF-8
    EscribirUTF8 rutaHTML, html

    ' 12. Exportar a PDF con Word
    ExportarWordAPDF rutaHTML, rutaPDF
End Sub


'Public Sub GenerarInformeHTML(rutaSalida As String, rutaPlantilla As String, _
'                              datos As Dictionary, archivosMD As Collection)
'
'    Dim html As String
'    Dim contenido As String
'    Dim frag As String
'    Dim rutaTemp As String
'    Dim rutaHTMLtemp As String
'    Dim md As Variant
'
'    ' 1. Cargar plantilla
'    html = LeerArchivoUTF8(rutaPlantilla)
'
'    ' 2. Sustituir marcadores de cabecera
'    Dim clave As Variant
'    For Each clave In datos.Keys
'        html = ReemplazarMarcadores(html, clave, datos(clave))
'    Next clave
'
'    ' 3. Convertir cada Markdown a HTML y concatenar
'    contenido = ""
'
'    For Each md In archivosMD
'        rutaTemp = md
'        rutaHTMLtemp = rutaTemp & ".html"
'
'        frag = ConvertirMarkdownAHTML(rutaTemp, rutaHTMLtemp)
'
'        ' Eliminar <html>, <body>, etc.
'        frag = Replace(frag, "<html>", "")
'        frag = Replace(frag, "</html>", "")
'        frag = Replace(frag, "<body>", "")
'        frag = Replace(frag, "</body>", "")
'        frag = Replace(frag, "<head>", "")
'        frag = Replace(frag, "</head>", "")
'
'        contenido = contenido & frag & vbCrLf
'    Next md
'
'    ' 4. Insertar contenido en la plantilla
'    html = ReemplazarMarcadores(html, "CONTENIDO", contenido)
'
'    ' 5. Guardar archivo final en UTF-8
'    EscribirUTF8 rutaSalida, html
'End Sub


Public Function LeerArchivoUTF8(Ruta As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    stm.Type = 2 'Texto
    stm.Charset = "UTF-8"
    stm.Open
    stm.LoadFromFile Ruta
    
    LeerArchivoUTF8 = stm.ReadText
    stm.Close
End Function


Public Function ReemplazarMarcadores(texto As String, clave As String, Valor As String) As String
    ReemplazarMarcadores = Replace(texto, "{{" & clave & "}}", Valor)
End Function

Public Function RellenarMarcadoresCiclos(html As String, _
                                         c1 As String, c2 As String, c3 As String, _
                                         r1 As String, r2 As String, r3 As String, _
                                         tc1 As String, tc2 As String, tc3 As String, _
                                         tl1 As String, tl2 As String, tl3 As String) As String

    ' --- Números de ciclo ---
    html = Replace(html, "{{CICLO1}}", c1)
    html = Replace(html, "{{CICLO2}}", c2)
    html = Replace(html, "{{CICLO3}}", c3)

    ' --- Rangos ---
    html = Replace(html, "{{CICLO1_RANGO}}", r1)
    html = Replace(html, "{{CICLO2_RANGO}}", r2)
    html = Replace(html, "{{CICLO3_RANGO}}", r3)

    ' --- Textos cortos ---
    html = Replace(html, "{{CICLO1_TEXTO_CORTO}}", tc1)
    html = Replace(html, "{{CICLO2_TEXTO_CORTO}}", tc2)
    html = Replace(html, "{{CICLO3_TEXTO_CORTO}}", tc3)

    ' --- Textos largos ---
    html = Replace(html, "{{CICLO1_TEXTO_LARGO}}", tl1)
    html = Replace(html, "{{CICLO2_TEXTO_LARGO}}", tl2)
    html = Replace(html, "{{CICLO3_TEXTO_LARGO}}", tl3)

    RellenarMarcadoresCiclos = html
End Function

Public Function RellenarMarcadoresPinaculosEscollos(html As String, _
                                                    P1 As String, P2 As String, P3 As String, P4 As String, _
                                                    E1 As String, E2 As String, E3 As String, E4 As String, _
                                                    TP1 As String, TP2 As String, TP3 As String, TP4 As String, _
                                                    TE1 As String, TE2 As String, TE3 As String, TE4 As String) As String

    ' --- Pináculos ---
    html = Replace(html, "{{P1}}", P1)
    html = Replace(html, "{{P2}}", P2)
    html = Replace(html, "{{P3}}", P3)
    html = Replace(html, "{{P4}}", P4)

    ' --- Escollos ---
    html = Replace(html, "{{E1}}", E1)
    html = Replace(html, "{{E2}}", E2)
    html = Replace(html, "{{E3}}", E3)
    html = Replace(html, "{{E4}}", E4)

    ' --- Textos cortos de Pináculos ---
    html = Replace(html, "{{TEXTO_P1}}", TP1)
    html = Replace(html, "{{TEXTO_P2}}", TP2)
    html = Replace(html, "{{TEXTO_P3}}", TP3)
    html = Replace(html, "{{TEXTO_P4}}", TP4)

    ' --- Textos largos de Escollos ---
    html = Replace(html, "{{TEXTO_E1}}", TE1)
    html = Replace(html, "{{TEXTO_E2}}", TE2)
    html = Replace(html, "{{TEXTO_E3}}", TE3)
    html = Replace(html, "{{TEXTO_E4}}", TE4)

    RellenarMarcadoresPinaculosEscollos = html
End Function

Public Function RellenarMarcadoresTransitos(html As String, _
                                            TF As String, _
                                            TM As String, _
                                            TE As String, _
                                            EA As String) As String

    ' --- Tránsitos actuales ---
    html = Replace(html, "{{TRANSITO_FISICO}}", TF)
    html = Replace(html, "{{TRANSITO_MENTAL}}", TM)
    html = Replace(html, "{{TRANSITO_ESPIRITUAL}}", TE)

'   +- Transito Físico
'   +- Transito Mental
'   +- Transito Emocional
'   +- Transito Espiritual

    ' --- Esencia anual ---
    html = Replace(html, "{{ESENCIA_ANUAL}}", EA)

    RellenarMarcadoresTransitos = html
End Function


Sub RellenaTablaTransitos()

Dim ED(1 To 10) As String
Dim TF(1 To 10) As String
Dim TM(1 To 10) As String
Dim TE(1 To 10) As String
Dim EDD(1 To 10) As String
Dim EDS(1 To 10) As String
Dim AP(1 To 10) As String

' Ejemplo:
For i = 1 To 10
    ED(i) = r.Edad(i)
    TF(i) = r.TransitoFisico(i)
    TM(i) = r.TransitoMental(i)
    TE(i) = r.TransitoEspiritual(i)
    EDD(i) = r.EsenciaDD(i)
    EDS(i) = r.EsenciaDS(i)
    AP(i) = r.AnioPersonal(i)
Next i

End Sub

Public Function RellenarMarcadoresProgresiones(html As String, _
                                               ED() As String, _
                                               TF() As String, _
                                               TM() As String, _
                                               TE() As String, _
                                               EDD() As String, _
                                               EDS() As String, _
                                               AP() As String) As String
    Dim i As Long

    ' --- Edades ---
    For i = 1 To UBound(ED)
        html = Replace(html, "{{ED" & i & "}}", ED(i))
    Next i

    ' --- Tránsito Físico ---
    For i = 1 To UBound(TF)
        html = Replace(html, "{{TF" & i & "}}", TF(i))
    Next i

    ' --- Tránsito Mental ---
    For i = 1 To UBound(TM)
        html = Replace(html, "{{TM" & i & "}}", TM(i))
    Next i

    ' --- Tránsito Espiritual ---
    For i = 1 To UBound(TE)
        html = Replace(html, "{{TE" & i & "}}", TE(i))
    Next i

    ' --- Esencia (dd) ---
    For i = 1 To UBound(EDD)
        html = Replace(html, "{{EDD" & i & "}}", EDD(i))
    Next i

    ' --- Esencia (ds) ---
    For i = 1 To UBound(EDS)
        html = Replace(html, "{{EDS" & i & "}}", EDS(i))
    Next i

    ' --- Año Personal ---
    For i = 1 To UBound(AP)
        html = Replace(html, "{{AP" & i & "}}", AP(i))
    Next i

    RellenarMarcadoresProgresiones = html
End Function

Sub RellenaCasas()

Dim c(1 To 9) As String
Dim p(1 To 9) As String
Dim M(1 To 9) As String
Dim CasaTexto(1 To 9) As String

For i = 1 To 9
    c(i) = r.CasaValor(i)
    p(i) = r.CasaPorcentaje(i)
    M(i) = r.CasaMedia(i)
    CasaTexto(i) = r.InterpretacionCasa(i)
Next i

End Sub


Public Function RellenarMarcadoresCasas(html As String, _
                                        c() As String, _
                                        p() As String, _
                                        M() As String, _
                                        CasaTexto() As String, _
                                        TextoValores As String, _
                                        TextoPorcentajes As String, _
                                        TextoMedias As String) As String
    Dim i As Long

    ' --- Valores ---
    For i = 1 To 9
        html = Replace(html, "{{C" & i & "}}", c(i))
    Next i

    ' --- Porcentajes ---
    For i = 1 To 9
        html = Replace(html, "{{P" & i & "}}", p(i))
    Next i

    ' --- Medias ---
    For i = 1 To 9
        html = Replace(html, "{{M" & i & "}}", M(i))
    Next i

    ' --- Interpretación por casa ---
    For i = 1 To 9
        html = Replace(html, "{{CASA" & i & "}}", CasaTexto(i))
    Next i

    ' --- Interpretaciones generales ---
    html = Replace(html, "{{TEXTO_VALORES}}", TextoValores)
    html = Replace(html, "{{TEXTO_PORCENTAJES}}", TextoPorcentajes)
    html = Replace(html, "{{TEXTO_MEDIAS}}", TextoMedias)

    RellenarMarcadoresCasas = html
End Function






Public Function ConvertirMarkdownAHTML(rutaMD As String, rutaHTML As String) As String
    Dim cmd As String
    cmd = "pandoc """ & rutaMD & """ -f markdown -t html -o """ & rutaHTML & """"
    Shell cmd, vbHide
    
    ' Esperar un poco a que termine
    Dim t As Single: t = Timer
    Do While Timer < t + 1
        DoEvents
    Loop
    
    ConvertirMarkdownAHTML = LeerArchivoUTF8(rutaHTML)
End Function

Public Sub ExportarWordAPDF(rutaHTML As String, rutaPDF As String)

    Dim wd As Object
    Set wd = CreateObject("Word.Application")
    
    wd.Visible = False
    wd.Documents.Open rutaHTML
    
    wd.ActiveDocument.ExportAsFixedFormat rutaPDF, 17 '17 = PDF
    
    wd.ActiveDocument.Close False
    wd.Quit

End Sub

