Attribute VB_Name = "modInterpreta"
Option Compare Database
Option Explicit

Sub interpreta(idPer As Long, idFon As Long, idRes As Long)

Dim strSQL As String

    Dim rs As DAO.Recordset

    strSQL = "SELECT IDPersona, IDFonetica, IDResultado, " & _
                    "NumeroAlma, NumeroDestino, NumeroPersonalidad, NumeroCaminoVida, " & _
                    "NumeroMadurez, AnioPersonal, EdadPersonal, " & _
                    "PlanoFisico, PlanoEmocional, PlanoMental, PlanoIntuitivo, " & _
                    "PiedraAngular, PiedraToque, " & _
                    "PrimeraLetra, PrimeraVocal, PrimeraConsonante, " & _
                    "RespuestaSubconsciente, Poder " & vbCrLf & _
             "FROM tbuResultados " & vbCrLf & _
             "WHERE IDPersona = " & idPer & " " & _
               "AND IDFonetica = " & idFon & " " & _
               "AND IDResultado = " & idRes & ";"


    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    
'Metadatos del cálculo (el contexto)
'Siempre primero.
'Sin contexto, los números no significan nada.

'Sistema de Cálculo
'Número de Ciclos
'Método de Ciclos
'Sistema de Tarot
'Versión del Motor
'Fecha del Cálculo
    
    
'Números principales (la identidad)

'Camino de Vida
'Destino / Expresión
'Alma
'Personalidad
'Día de Nacimiento


'Números derivados (la evolución interna)

'Madurez
'Poder
'Respuesta Subconsciente


'Planos de expresión (cómo se manifiesta)

'Plano Físico
'Plano Emocional
'Plano Mental
'Plano Intuitivo


'Árbol de Vida (si lo usas en la interpretación)

'Árbol Paterno
'Árbol Materno
'Árbol de Vida


'Letras y símbolos (la vibración inicial)
'Son detalles finos pero muy reveladores.

'primera letra
'primera Vocal
'primera Consonante
'Piedra Angular
'Piedra de Toque


'Temporales (el momento actual)

'Año Personal
'   mes Personal    (no lo tengo)
'   Día Personal    (no lo tengo)
'Edad Personal
'Esencia


'Listas especiales
'Para completar la lectura.

'Números Ausentes
'Números Dominantes



End Sub


