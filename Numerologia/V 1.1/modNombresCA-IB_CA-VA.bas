Attribute VB_Name = "modNombresCA-IB_CA-VA"

Option Compare Database
Option Explicit

' ============================================================================
'  MÓDULO TEMPORAL PARA CARGAR EL DICCIONARIO BALEAR (CA-IB)
'  Ejecutar una sola vez:  CargarDiccionarioBalear
'  Luego exportar y eliminar el módulo de la aplicación
' ============================================================================

' ============================================================================
'  PROCEDIMIENTO MAESTRO
' ============================================================================

Public Sub CargarDiccionarioBalear()

    Call CargarNombresBalear
    Call CargarApellidosBalear

    MsgBox "Diccionario balear (CA-IB) cargado correctamente.", vbInformation

End Sub


' ============================================================================
'  FUNCIÓN AUXILIAR PARA INSERTAR EN TABLAS
' ============================================================================

'Private Sub InsertarEntradaDiccionario( _
'    ByVal tabla As String, _
'    ByVal palabra As String, _
'    ByVal fonema As String, _
'    ByVal idioma As String, _
'    ByVal tipo As String, _
'    Optional ByVal notas As String = "" _
'    )
'
'    Dim sql As String
'
'    sql = "INSERT INTO " & tabla & " (Palabra, FonemaCon, Idioma, TipoEntrada, Notas, Activo) " & _
'          "VALUES (" & _
'          "'" & Replace(UCase$(palabra), "'", "''") & "', " & _
'          "'" & Replace(fonema, "'", "''") & "', " & _
'          "'" & idioma & "', " & _
'          "'" & tipo & "', " & _
'          "'" & Replace(notas, "'", "''") & "', " & _
'          "True)"
'
'    CurrentDb.Execute sql, dbFailOnError
'
'End Sub



Public Sub PoblarNombresCA_IB()


' === MASCULINS — FORMES CANÒNIQUES ===

AgregarEntradaDiccionario "CA-IB", "JAUME", "Yauma", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "JOAN", "Yoan", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MATEU", "Mateu", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MIQUEL", "Miquel", "NOMBRE", "Català comú en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "GABRIEL", "Gabriel", "NOMBRE", "Tradicional bíblic", "Balears"
AgregarEntradaDiccionario "CA-IB", "RAFAEL", "Rafel", "NOMBRE", "Tradicional bíblic", "Balears"
AgregarEntradaDiccionario "CA-IB", "TOMÀS", "Tomas", "NOMBRE", "Tradicional bíblic", "Balears"
AgregarEntradaDiccionario "CA-IB", "JERONI", "Yeroni", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ELIES", "Elies", "NOMBRE", "Tradicional bíblic", "Balears"
AgregarEntradaDiccionario "CA-IB", "BARTOMEU", "Bartomeu", "NOMBRE", "Forma balear tradicional", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "GREGORI", "Gregori", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "COSME", "Cosme", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "SEVER", "Sever", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "APOL·LONI", "Apolloni", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BENEDET", "Benedet", "NOMBRE", "Tradicional balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "BERNAT", "Bernat", "NOMBRE", "Català comú en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "LLUC", "Lluc", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ARNAU", "Arnau", "NOMBRE", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "JAN", "Yan", "NOMBRE", "Ús modern amb fonètica balear", "Balears"

AgregarEntradaDiccionario "CA-IB", "ANDREU", "Andreu", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "FRANCESC", "Francesc", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "PERE", "Pere", "NOMBRE", "Forma catalana en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "SALVADOR", "Salvador", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "CRISTÒFOR", "Cristofor", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "NICOLAU", "Nicolau", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "FERRAN", "Ferran", "NOMBRE", "Tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "JAUME JOAN", "Yaume Yoan", "NOMBRE", "Compost tradicional balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "ONOFRE", "Onofre", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "RAFELET", "Rafelet", "NOMBRE", "Tradicional antic", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "RAFEU", "Rafeu", "NOMBRE", "Variant tradicional de Rafel", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ISIDRE", "Isidre", "NOMBRE", "Tradicional català en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ESTEVE", "Esteve", "NOMBRE", "Tradicional català en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "FERRAN", "Ferran", "NOMBRE", "Tradicional català en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "JEREMIES", "Yeremies", "NOMBRE", "Tradicional bíblic", "Balears"
'AgregarEntradaDiccionario "CA-IB", "CRISTÒFOR", "Cristofor", "NOMBRE", "Tradicional balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "SALVADOR", "Salvador", "NOMBRE", "Tradicional balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "ANDREU", "Andreu", "NOMBRE", "Tradicional balear", "Balears"

' === FEMENINS — FORMES CANÒNIQUES ===

AgregarEntradaDiccionario "CA-IB", "CATALINA", "Catalina", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MARGALIDA", "Margalida", "NOMBRE", "Forma balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "JOANA", "Yoana", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MIQUELA", "Miquela", "NOMBRE", "Forma balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "ANTÒNIA", "Antonia", "NOMBRE", "Forma balear tradicional", "Balears"

AgregarEntradaDiccionario "CA-IB", "APOL·LÒNIA", "Apollonia", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "GREGÒRIA", "Gregoria", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "SEVERA", "Severa", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "SEVERINA", "Severina", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BENEDETA", "Benedeta", "NOMBRE", "Tradicional balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "LLÚCIA", "Llúcia", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "EULÀLIA", "Eulalia", "NOMBRE", "Tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "NEUS", "Neus", "NOMBRE", "Català comú en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "NÚRIA", "Nuria", "NOMBRE", "Català comú en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "CARME", "Carme", "NOMBRE", "Català comú en ús balear", "Balears"

AgregarEntradaDiccionario "CA-IB", "MAGDALENA", "Magdalena", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "FRANCINA", "Francina", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ELIONOR", "Elionor", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "CLARISSA", "Clarissa", "NOMBRE", "Tradicional antic", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "CATALINETA", "Catalineta", "NOMBRE", "Forma tradicional documentada", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "MARGARITA", "Margarita", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "EULÀLIA", "Eulalia", "NOMBRE", "Tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "ADELINA", "Adelina", "NOMBRE", "Tradicional antic", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "AINA", "Aina", "NOMBRE", "Forma tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "COLOMA", "Coloma", "NOMBRE", "Tradicional balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "FRANCISCA", "Francisca", "NOMBRE", "Forma canònica tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "ELIONOR", "Elionor", "NOMBRE", "Tradicional balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ÚRSULA", "Ursula", "NOMBRE", "Tradicional antic", "Balears"
AgregarEntradaDiccionario "CA-IB", "TECLA", "Tecla", "NOMBRE", "Tradicional català en ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ELVIRA", "Elvira", "NOMBRE", "Tradicional antic", "Balears"
AgregarEntradaDiccionario "CA-IB", "CLARISSA", "Clarissa", "NOMBRE", "Tradicional antic", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ADELINA", "Adelina", "NOMBRE", "Tradicional antic", "Mallorca"

' === HIPOCORÍSTICS — DOCUMENTATS ===

AgregarEntradaDiccionario "CA-IB", "TONI", "Toni", "NOMBRE", "Hipocorístic balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "TONA", "Tona", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BIEL", "Biel", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "TOMEU", "Tomeu", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "TÒFOL", "Tofol", "NOMBRE", "Hipocorístic balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "GORI", "Gori", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "XISCA", "Shisca", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "XISCO", "Shisco", "NOMBRE", "Hipocorístic balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "CATI", "Cati", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "XIM", "Shim", "NOMBRE", "Hipocorístic balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "XIMA", "Shima", "NOMBRE", "Hipocorístic balear", "Mallorca"

AgregarEntradaDiccionario "CA-IB", "NOFRE", "Nofre", "NOMBRE", "Hipocorístic balear d'Onofre", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "TOLO", "Tolo", "NOMBRE", "Hipocorístic balear d'Antoni", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "PACO", "Paco", "NOMBRE", "Hipocorístic tradicional d'en Francesc", "Balears"
'AgregarEntradaDiccionario "CA-IB", "XIMET", "Ximet", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "XIMETA", "Ximeta", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "XIMONA", "Ximona", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"

'AgregarEntradaDiccionario "CA-IB", "NOFRE", "Nofre", "NOMBRE", "Hipocorístic balear d'Onofre", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "TOLO", "Tolo", "NOMBRE", "Hipocorístic balear d'Antoni", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "PACO", "Paco", "NOMBRE", "Hipocorístic tradicional d'en Francesc", "Balears"

' === VARIANTS DIALECTALS ===

AgregarEntradaDiccionario "CA-IB", "GUIEM", "Guiem", "NOMBRE", "Variant balear de Guillem", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "YOANA", "Yoana", "NOMBRE", "Variant fonètica balear de Joana", "Balears"
AgregarEntradaDiccionario "CA-IB", "JAUME", "Yauma", "NOMBRE", "Variant fonètica balear de Jaume", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YORDI", "Yordi", "NOMBRE", "Variant fonètica balear de Jordi", "Balears"
AgregarEntradaDiccionario "CA-IB", "JULIA", "Yulia", "NOMBRE", "Variant balear de Júlia", "Balears"

'AgregarEntradaDiccionario "CA-IB", "YOAN", "Yoan", "NOMBRE", "Variant fonètica balear de Joan", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YERONI", "Yeroni", "NOMBRE", "Variant fonètica balear de Jeroni", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "YELIES", "Yelies", "NOMBRE", "Variant fonètica balear d'Elies", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "YANA", "Yana", "NOMBRE", "Variant fonètica balear de Jana", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YULIA", "Yulia", "NOMBRE", "Variant fonètica balear de Júlia", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YORDI", "Yordi", "NOMBRE", "Variant fonètica balear de Jordi", "Balears"
'AgregarEntradaDiccionario "CA-IB", "GUIU", "Guiu", "NOMBRE", "Variant balear de Guillem antic", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "GUIEMO", "Guiemo", "NOMBRE", "Variant fonètica antiga de Guiem", "Mallorca"

'AgregarEntradaDiccionario "CA-IB", "YOAN", "Yoan", "NOMBRE", "Variant fonètica balear de Joan", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YERONI", "Yeroni", "NOMBRE", "Variant fonètica balear de Jeroni", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "YELIES", "Yelies", "NOMBRE", "Variant fonètica balear d'Elies", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "YORDI", "Yordi", "NOMBRE", "Variant fonètica balear de Jordi", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YULIA", "Yulia", "NOMBRE", "Variant balear de Júlia", "Balears"
'AgregarEntradaDiccionario "CA-IB", "YANA", "Yana", "NOMBRE", "Variant fonètica balear de Jana", "Balears"
'AgregarEntradaDiccionario "CA-IB", "GUIU", "Guiu", "NOMBRE", "Variant antiga de Guillem", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "GUIEMO", "Guiemo", "NOMBRE", "Variant fonètica antiga de Guiem", "Mallorca"


End Sub

Private Sub CargarNombresValenciano()

' === MASCULINS — FORMES CANÒNIQUES CA-VA ===

AgregarEntradaDiccionario "CA-VA", "VICENT", "Visent", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "JOSEP", "Yosep", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "JOAN", "Yoan", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "PERE", "Pere", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "FRANCESC", "Francesc", "NOMBRE", "Forma valenciana tradicional", "València"

AgregarEntradaDiccionario "CA-VA", "ENRIC", "Enric", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "FERRAN", "Ferran", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ANDREU", "Andreu", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "MATEU", "Mateu", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "MIQUEL", "Miquel", "NOMBRE", "Forma valenciana tradicional", "València"

AgregarEntradaDiccionario "CA-VA", "RAFAEL", "Rafael", "NOMBRE", "Forma bíblica tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "GABRIEL", "Gabriel", "NOMBRE", "Forma bíblica tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "TOMÀS", "Tomas", "NOMBRE", "Forma bíblica tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ELIES", "Elies", "NOMBRE", "Forma bíblica tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ISIDRE", "Isidre", "NOMBRE", "Forma tradicional valenciana", "València"

'--------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "BERNAT", "Bernat", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GUILLEM", "Guillem", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "NICOLAU", "Nicolau", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CRISTÒFOR", "Cristofor", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "SEBASTIÀ", "Sebastia", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "ROGER", "Roger", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ADRIÀ", "Adria", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GENÍS", "Genis", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "JORDI", "Yordi", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "XAVIER", "Shavier", "NOMBRE", "Tradicional valencià", "València"

'--------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "MARC", "Marc", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ÀLEX", "Alex", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ERIC", "Eric", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BRUNO", "Bruno", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "IAN", "Yan", "NOMBRE", "Ús modern valencià", "València"

AgregarEntradaDiccionario "CA-VA", "PAU", "Pau", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ARNAU", "Arnau", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GUILLEM", "Guillem", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "TEO", "Teo", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "LUCAS", "Lucas", "NOMBRE", "Ús modern valencià", "València"

'-----------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "AGUSTÍ", "Agusti", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ALFONS", "Alfons", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AMBRÒS", "Ambros", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ARNAU", "Arnau", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BALTAZAR", "Baltazar", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "BENET", "Benet", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BLASCO", "Blasco", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BONAVENTURA", "Bonaventura", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "DOMÈNEC", "Domenec", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ESTEVE", "Esteve", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "FELIU", "Feliu", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GENÍS", "Genis", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GREGORI", "Gregori", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "HILARI", "Hilari", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ISMAEL", "Ismael", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "JERONI", "Yeroni", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "LLORENÇ", "Llorenc", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MANEL", "Manel", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MAURI", "Mauri", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "NARCÍS", "Narcis", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "ODÓ", "Odo", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "PASCUAL", "Pascual", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "SALVADOR", "Salvador", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "SIMÓ", "Simo", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "TOMÀS", "Tomas", "NOMBRE", "Tradicional valencià", "València"


' === FEMENINS — FORMES CANÒNIQUES CA-VA ===

AgregarEntradaDiccionario "CA-VA", "TERESA", "Teresa", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ISABEL", "Isabel", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ANNA", "Anna", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "CATERINA", "Caterina", "NOMBRE", "Forma valenciana tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "MARIA", "Maria", "NOMBRE", "Forma tradicional", "València"

AgregarEntradaDiccionario "CA-VA", "FRANCISCA", "Francisca", "NOMBRE", "Forma tradicional valenciana", "València"
AgregarEntradaDiccionario "CA-VA", "EULÀLIA", "Eulalia", "NOMBRE", "Forma tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "CLARA", "Clara", "NOMBRE", "Forma tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "MARTA", "Marta", "NOMBRE", "Forma tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "NÚRIA", "Nuria", "NOMBRE", "Forma tradicional", "València"

'--------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "ADELINA", "Adelina", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AGUEDA", "Agueda", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AMÀLIA", "Amalia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ANGELA", "Angela", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CARME", "Carme", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "DOLORS", "Dolors", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "HORTÈNSIA", "Hortensia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MERCÈ", "Merce", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ROSER", "Roser", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "VERÒNICA", "Veronica", "NOMBRE", "Tradicional valencià", "València"

'--------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "JÚLIA", "Julia", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MIREIA", "Mireia", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AINA", "Aina", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "NOA", "Noa", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "LIA", "Lia", "NOMBRE", "Ús modern valencià", "València"

AgregarEntradaDiccionario "CA-VA", "ARLET", "Arlet", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "GALA", "Gala", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ONA", "Ona", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "IRIS", "Iris", "NOMBRE", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "NORA", "Nora", "NOMBRE", "Ús modern valencià", "València"

'--------------------------------------------------------

AgregarEntradaDiccionario "CA-VA", "AGUEDA", "Agueda", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AMÀLIA", "Amalia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ANGELA", "Angela", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ANTÒNIA", "Antonia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "AURORA", "Aurora", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "BÀRBARA", "Barbara", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BEATRIU", "Beatriu", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BERNARDA", "Bernarda", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CARLOTA", "Carlota", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CARMINA", "Carmina", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "CONSOL", "Consol", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "DOLORS", "Dolors", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "EUGÈNIA", "Eugenia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "HORTÈNSIA", "Hortensia", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MAGDALENA", "Magdalena", "NOMBRE", "Tradicional valencià", "València"

AgregarEntradaDiccionario "CA-VA", "MARGARIDA", "Margarida", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MERCÈ", "Merce", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "PAULA", "Paula", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "RAQUEL", "Raquel", "NOMBRE", "Tradicional valencià", "València"
AgregarEntradaDiccionario "CA-VA", "VERÒNICA", "Veronica", "NOMBRE", "Tradicional valencià", "València"

' === HIPOCORÍSTICS — DOCUMENTATS CA-VA ===

AgregarEntradaDiccionario "CA-VA", "VICENTET", "Vicentet", "NOMBRE", "Hipocorístic tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "XIMO", "Shimo", "NOMBRE", "Hipocorístic de Joaquim", "València"
AgregarEntradaDiccionario "CA-VA", "QUIM", "Quim", "NOMBRE", "Hipocorístic de Joaquim", "València"
AgregarEntradaDiccionario "CA-VA", "PACO", "Paco", "NOMBRE", "Hipocorístic tradicional de Francesc", "València"
AgregarEntradaDiccionario "CA-VA", "PEP", "Pep", "NOMBRE", "Hipocorístic de Josep", "València"

AgregarEntradaDiccionario "CA-VA", "PERELO", "Perelo", "NOMBRE", "Hipocorístic tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "TONI", "Toni", "NOMBRE", "Hipocorístic d'Antoni", "València"
AgregarEntradaDiccionario "CA-VA", "TONET", "Tonet", "NOMBRE", "Hipocorístic tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "MARIAETA", "Mariaeta", "NOMBRE", "Hipocorístic tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "CONXA", "Conxa", "NOMBRE", "Hipocorístic d'Assumpció", "València"


' === VARIANTS DIALECTALS CA-VA ===

AgregarEntradaDiccionario "CA-VA", "YOAN", "Yoan", "NOMBRE", "Variant fonètica valenciana de Joan", "València"
AgregarEntradaDiccionario "CA-VA", "YOSEP", "Yosep", "NOMBRE", "Variant fonètica valenciana de Josep", "València"
AgregarEntradaDiccionario "CA-VA", "VISENT", "Visent", "NOMBRE", "Variant fonètica valenciana de Vicent", "València"
AgregarEntradaDiccionario "CA-VA", "YOAQUIM", "Yoaquim", "NOMBRE", "Variant fonètica valenciana de Joaquim", "València"
AgregarEntradaDiccionario "CA-VA", "YOANA", "Yoana", "NOMBRE", "Variant fonètica valenciana de Joana", "València"


End Sub


'    ' === BLOQUE 1 (1–50) ===
'
'    AgregarEntradaDiccionario "CA-IB", "AINA", "AINA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MARGALIDA", "MARGALIDA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "COLOMA", "COLOMA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLÚCIA", "LLUCIA", "NOMBRE", "Variant balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUCIA", "LLUCIA", "NOMBRE", "Variant balear", "Mallorca / Menorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUCÍA", "LLUCIA", "NOMBRE", "Variant balear", "Menorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUCIÁ", "LLUCIA", "NOMBRE", "Equivalent balear de Luciano", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUC", "LLUC", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BIEL", "BIEL", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TOMEU", "TOMEU", "NOMBRE", "Balear", "Mallorca"
'
'    AgregarEntradaDiccionario "CA-IB", "RAFEL", "RAFEL", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "CATALINA", "CATALINA", "NOMBRE", "Balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "CATI", "CATI", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "JOANA", "YOANA", "NOMBRE", "Balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "BARTOMEU", "BARTOMEU", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MATEU", "MATEU", "NOMBRE", "Balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MARIONA", "MARIONA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MIQUELO", "MIQUELO", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MIQUELA", "MIQUELA", "NOMBRE", "Balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MIQUELETA", "MIQUELETA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'
'    AgregarEntradaDiccionario "CA-IB", "NOFRE", "NOFRE", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "ONOFRE", "ONOFRE", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "GORI", "GORI", "NOMBRE", "Hipocorístic balear", "Mallorca"
'
'    AgregarEntradaDiccionario "CA-IB", "XISCA", "SHISCA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XISCO", "SHISCO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TÒFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TÓFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TOLO", "TOLO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'
'    AgregarEntradaDiccionario "CA-IB", "PAULETA", "PAULETA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "PAULET", "PAULET", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIM", "SHIM", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMA", "SHIMA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMET", "SHIMET", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMETA", "SHIMETA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMONA", "SHIMONA", "NOMBRE", "Balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMOT", "SHIMOT", "NOMBRE", "Balear", "Mallorca"
'
'    ' === Nombres añadidos para completar el bloque ===
'
'    AgregarEntradaDiccionario "CA-IB", "GABRIELA", "GABRIELA", "NOMBRE", "Forma balear moderna", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "GABRIEL", "GABRIEL", "NOMBRE", "Forma balear moderna", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NURIA", "NURIA", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NEUS", "NEUS", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "CARME", "CARME", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ANTÒNIA", "ANTONIA", "NOMBRE", "Forma moderna balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ANTONI", "ANTONI", "NOMBRE", "Forma moderna balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JAUME", "YAUME", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "PERE", "PERE", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MARIA", "MARIA", "NOMBRE", "Català comú en ús balear", "Balears"
'
'
'    ' === BLOQUE 2 (51–100) — VERSIÓ DEFINITIVA ===
'
'    ' --- Tradicionals i d’ús general ---
'    AgregarEntradaDiccionario "CA-IB", "JAUME", "YAUME", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ANTONI", "ANTONI", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ANTÒNIA", "ANTONIA", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "PERE", "PERE", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "PAU", "PAU", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "CARME", "CARME", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NEUS", "NEUS", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NÚRIA", "NURIA", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "GUILLEM", "GUILLEM", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "GUIEM", "GUIEM", "NOMBRE", "Variant balear", "Mallorca"
'
'    ' --- Hipocorístics balears documentats (autèntics) ---
'    AgregarEntradaDiccionario "CA-IB", "TONI", "TONI", "NOMBRE", "Hipocorístic balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "TONA", "TONA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MIQUELO", "MIQUELO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BIEL", "BIEL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TOMEU", "TOMEU", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TÒFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "GORI", "GORI", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "PERICO", "PERICO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XISCA", "SHISCA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XISCO", "SHISCO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'
'    ' --- Moderns d’ús real a Balears ---
'    AgregarEntradaDiccionario "CA-IB", "LAURA", "LAURA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "CLARA", "CLARA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MARTA", "MARTA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JÚLIA", "JULIA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ARNAU", "ARNAU", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "BRUNA", "BRUNA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "BERTA", "BERTA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "CARLA", "CARLA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MARC", "MARC", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JORDI", "YORDI", "NOMBRE", "Ús general a Balears", "Balears"
'
'    ' --- Tradicionals menys freqüents però 100% CA-IB ---
'    AgregarEntradaDiccionario "CA-IB", "EULÀLIA", "EULALIA", "NOMBRE", "Tradicional", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "EULALI", "EULALI", "NOMBRE", "Tradicional", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "APOL·LÒNIA", "APOLLONIA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "APOL·LONI", "APOLLONI", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "COSME", "COSME", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMET", "DAMET", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMETA", "DAMETA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVER", "SEVER", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERA", "SEVERA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERINA", "SEVERINA", "NOMBRE", "Tradicional balear", "Mallorca"
'
'
'    ' === BLOQUE 3 (101–150) — VERSIÓ DEFINITIVA ===
'
'    ' --- Tradicionals CA-IB ---
'    AgregarEntradaDiccionario "CA-IB", "BENEDET", "BENEDET", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BENEDETA", "BENEDETA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BERNAT", "BERNAT", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "BERNADA", "BERNADA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MARGALIDA", "MARGALIDA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "CATALINA", "CATALINA", "NOMBRE", "Tradicional balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JOANA", "YOANA", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MATEU", "MATEU", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MIQUELA", "MIQUELA", "NOMBRE", "Forma balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "MIQUEL", "MIQUEL", "NOMBRE", "Català comú en ús balear", "Balears"
'
'    ' --- Hipocorístics balears (autèntics i documentats) ---
'    AgregarEntradaDiccionario "CA-IB", "TOMEUET", "TOMEUET", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "CATI", "CATI", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIM", "SHIM", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMA", "SHIMA", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMET", "SHIMET", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMETA", "SHIMETA", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIMONA", "SHIMONA", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TOLO", "TOLO", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "NOFRE", "NOFRE", "NOMBRE", "Hipocorístic balear documentat", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "ONOFRE", "ONOFRE", "NOMBRE", "Forma balear", "Mallorca"
'
'    ' --- Moderns d’ús real a Balears ---
'    AgregarEntradaDiccionario "CA-IB", "AINA", "AINA", "NOMBRE", "Ús general a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NEREA", "NEREA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "IRIS", "IRIS", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NIL", "NIL", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JAN", "YAN", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ONA", "ONA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "POL", "POL", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ERIC", "ERIC", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "GALA", "GALA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "IAN", "YAN", "NOMBRE", "Ús modern a Balears", "Balears"
'
'    ' --- Tradicionals menys freqüents però 100% CA-IB ---
'    AgregarEntradaDiccionario "CA-IB", "APOL·LÒNIA", "APOLLONIA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "APOL·LONI", "APOLLONI", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "COSME", "COSME", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMET", "DAMET", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMETA", "DAMETA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVER", "SEVER", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERA", "SEVERA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERINA", "SEVERINA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUC", "LLUC", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLÚCIA", "LLUCIA", "NOMBRE", "Tradicional balear", "Mallorca"
'
'
'    ' === BLOQUE 4 (151–200) — VERSIÓ DEFINITIVA ===
'
'    ' --- Tradicionals CA-IB ---
'    AgregarEntradaDiccionario "CA-IB", "FELIU", "FELIU", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "GREGORI", "GREGORI", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "GREGÒRIA", "GREGORIA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BENEDET", "BENEDET", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BENEDETA", "BENEDETA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BERNAT", "BERNAT", "NOMBRE", "Català comú en ús balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "BERNADA", "BERNADA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "MARGALIDA", "MARGALIDA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "CATALINA", "CATALINA", "NOMBRE", "Tradicional balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "JOANA", "YOANA", "NOMBRE", "Forma balear", "Balears"
'
'    ' --- Hipocorístics CA-IB (autèntics i documentats) ---
'    AgregarEntradaDiccionario "CA-IB", "TONI", "TONI", "NOMBRE", "Hipocorístic balear", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "TONA", "TONA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "BIEL", "BIEL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TOMEU", "TOMEU", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "TÒFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "GORI", "GORI", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "PERICO", "PERICO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XISCA", "SHISCA", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XISCO", "SHISCO", "NOMBRE", "Hipocorístic balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "XIM", "SHIM", "NOMBRE", "Hipocorístic balear", "Mallorca"
'
'    ' --- Moderns d’ús real a Balears ---
'    AgregarEntradaDiccionario "CA-IB", "ARLET", "ARLET", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ELNA", "ELNA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "TEO", "TEO", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "TEIA", "TEIA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "IAN", "YAN", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "GALA", "GALA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "NIL", "NIL", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ONA", "ONA", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "POL", "POL", "NOMBRE", "Ús modern a Balears", "Balears"
'    AgregarEntradaDiccionario "CA-IB", "ERIC", "ERIC", "NOMBRE", "Ús modern a Balears", "Balears"
'
'    ' --- Tradicionals menys freqüents però 100% CA-IB ---
'    AgregarEntradaDiccionario "CA-IB", "APOL·LÒNIA", "APOLLONIA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "APOL·LONI", "APOLLONI", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "COSME", "COSME", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMET", "DAMET", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "DAMETA", "DAMETA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVER", "SEVER", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERA", "SEVERA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "SEVERINA", "SEVERINA", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLUC", "LLUC", "NOMBRE", "Tradicional balear", "Mallorca"
'    AgregarEntradaDiccionario "CA-IB", "LLÚCIA", "LLUCIA", "NOMBRE", "Tradicional balear", "Mallorca"

'End Sub


'Public Sub PoblarNombresCA_IB()
'
'    ' ==== NOMS BALEARS (CA-IB) ====
'    ' Exclusius, hipocorístics amb entitat pròpia
'    ' i catalans comuns amb accentuació balear.
'    ' Sense compostos, sense diminutius productius.
'
'    ' ==== BLOQUE 1 – NOMS EXCLUSIUS BALEARS (1–50) ====
'
'    ' --- Femenins exclusius ---
'    AgregarEntradaDiccionario "CA-IB", "AINA", "AINA", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MARGALIDA", "MARGALIDA", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "COLOMA", "COLOMA", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "LLUCIA", "LLUCIA", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "LLÚCIA", "LLUCIA", "NOMBRE", "Balear", ""
'
'    ' --- Masculins exclusius ---
'    AgregarEntradaDiccionario "CA-IB", "BIEL", "BIEL", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "TOMEU", "TOMEU", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "RAFEL", "RAFEL", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "LLUC", "LLUC", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MIQUELÓ", "MIQUELO", "NOMBRE", "Balear", ""
'
'    ' --- Hipocorístics amb entitat pròpia ---
'    AgregarEntradaDiccionario "CA-IB", "XISCA", "SHISCA", "NOMBRE", "Hipocorístic balear", ""
'    AgregarEntradaDiccionario "CA-IB", "XISCO", "SHISCO", "NOMBRE", "Hipocorístic balear", ""
'    AgregarEntradaDiccionario "CA-IB", "TÒFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", ""
'    AgregarEntradaDiccionario "CA-IB", "TÓFOL", "TOFOL", "NOMBRE", "Hipocorístic balear", ""
'    AgregarEntradaDiccionario "CA-IB", "TOLO", "TOLO", "NOMBRE", "Hipocorístic balear", ""
'    AgregarEntradaDiccionario "CA-IB", "NOFRE", "NOFRE", "NOMBRE", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "JOANET", "YOANET", "NOMBRE", "Hipocorístic balear", ""
'
'    ' (Puedes seguir ampliando aquí con más exclusius documentats)
'
'    Debug.Print "PoblarNombresCA_IB – Bloque 1 completado."
'
'
'    ' ==== BLOQUE 2 – CATALANS AMB ACCENTUACIÓ BALEAR (51–100) ====
'
'    ' --- Femenins amb accent balear ---
'    AgregarEntradaDiccionario "CA-IB", "ANTÒNIA", "ANTONIA", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "ANTÓNIA", "ANTONIA", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MÒNICA", "MONICA", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MÓNICA", "MONICA", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MIQUELA", "MIQUELA", "NOMBRE", "Català balear", ""
'
'    ' --- Masculins amb accent balear ---
'    AgregarEntradaDiccionario "CA-IB", "RAFÈL", "RAFEL", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "RAFÉL", "RAFEL", "NOMBRE", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "JAUMÉ", "JAUME", "NOMBRE", "Català balear", ""
'
'    ' (Aquí puedes añadir más variants accentuades pròpies de Balears)
'
'    Debug.Print "PoblarNombresCA_IB – Bloque 2 completado."
'
'
'    ' ==== BLOQUE 3 – CATALANS COMUNS (SUPORT CA-IB) (101–150) ====
'    ' Només alguns noms catalans comuns que es volen marcar explícitament
'    ' com a presents en CA-IB, tot i que ja existeixen en CA.
'
'    Dim comuns As Variant
'    Dim i As Long
'
'    comuns = Array( _
'        "JOAN", "PERE", "MIQUEL", "JAUME", "FRANCESC", "JOSEP", "MARIA", _
'        "CLARA", "PAULA", "CARME", "MONTSE", "ROSA", "TERESA", _
'        "BERNAT", "GUILLEM", "ANDREU", "MARTÍ", "PAU", "ARNAU", _
'        "ADRIÀ", "MARINA", "LAIA", "JÚLIA", "NEUS", "DOLORS" _
'    )
'
'    For i = LBound(comuns) To UBound(comuns)
'        ' Fonètica simplificada: acents fora, J inicial ? Y, X inicial ? SH si cal.
'        Select Case comuns(i)
'            Case "JOAN"
'                AgregarEntradaDiccionario "CA-IB", "JOAN", "YOAN", "NOMBRE", "Català comú CA-IB", ""
'            Case Else
'                AgregarEntradaDiccionario "CA-IB", comuns(i), _
'                                          Replace(comuns(i), "Í", "I"), _
'                                          "NOMBRE", "Català comú CA-IB", ""
'        End Select
'    Next i
'
'    Debug.Print "PoblarNombresCA_IB – Bloque 3 completado."
'    Debug.Print "Fin PoblarNombresCA_IB."
'
'End Sub
'
'
'
'Public Sub PoblarApellidosCA_IB()
'
'    ' ==== COGNOMS BALEARS (CA-IB) ====
'    ' Exclusius de Balears i catalans amb accentuació balear.
'    ' Sense compostos artificials.
'
'    ' ==== BLOQUE 1 – COGNOMS EXCLUSIUS BALEARS (1–50) ====
'
'    AgregarEntradaDiccionario "CA-IB", "PONS", "PONS", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "CANYELLES", "CANYELLES", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "LLABRÉS", "LLABRES", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "BENNÀSSAR", "BENNASSAR", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "ALOMAR", "ALOMAR", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "SUREDA", "SUREDA", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "SOCIÁS", "SOCIAS", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "GARAU", "GARAU", "APELLIDO", "Balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MASCARÓ", "MASCARO", "APELLIDO", "Balear", ""
'
'    ' (Aquí puedes añadir más cognoms exclusius de Balears)
'
'    Debug.Print "PoblarApellidosCA_IB – Bloque 1 completado."
'
'
'    ' ==== BLOQUE 2 – COGNOMS CATALANS AMB ACCENTUACIÓ BALEAR (51–100) ====
'
'    AgregarEntradaDiccionario "CA-IB", "FORNÉS", "FORNES", "APELLIDO", "Català balear", ""
'    AgregarEntradaDiccionario "CA-IB", "MIRÓ", "MIRO", "APELLIDO", "Català balear", ""
'
'    ' (Afegir aquí altres cognoms catalans amb accent balear si els vols marcar)
'
'    Debug.Print "PoblarApellidosCA_IB – Bloque 2 completado."
'    Debug.Print "Fin PoblarApellidosCA_IB."
'
'End Sub
'
'
'
'
'
'' ============================================================================
''  NOMBRES BALEARES (CA-IB)
'' ============================================================================
'
'Private Sub CargarNombresBalear()
'
'    ' -------------------------------
'    ' Exclusivos de Baleares
'    ' -------------------------------
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Aina", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Biel", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Tomeu", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Rafel", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Margalida", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Xisca", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Tòfol", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Tófol", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Miqueló", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Coloma", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Lluc", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Llucia", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Llúcia", "", "CA-IB", "NOMBRE"
'
'    ' -------------------------------
'    ' Hipocorísticos baleares
'    ' -------------------------------
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Joanet", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Nofre", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Tolo", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Miquelet", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Xisco", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Xisquet", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Xisqueta", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Pepet", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Pep", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Pepa", "", "CA-IB", "NOMBRE"
'
'    ' -------------------------------
'    ' Catalanes comunes con acentuación balear
'    ' -------------------------------
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Salvá", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Antònia", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Antónia", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Mònica", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Mónica", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Rafèl", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Rafél", "", "CA-IB", "NOMBRE"
'    InsertarEntradaDiccionario "tbmDicFonemasNom", "Jaumé", "", "CA-IB", "NOMBRE"
'
'    ' -------------------------------
'    ' Catalanes comunes (para completar CA-IB)
'    ' -------------------------------
'    Dim comunes As Variant
'    comunes = Array( _
'        "Joan", "Pere", "Miquel", "Jaume", "Francesc", "Josep", "Maria", _
'        "Clara", "Paula", "Carme", "Montserrat", "Rosa", "Teresa", _
'        "Bernat", "Guillem", "Andreu", "Martí", "Pau", "Arnau", _
'        "Adrià", "Marina", "Laia", "Júlia", "Neus", "Dolors" _
'    )
'
'    Dim i As Long
'    For i = LBound(comunes) To UBound(comunes)
'        InsertarEntradaDiccionario "tbmDicFonemasNom", comunes(i), "", "CA-IB", "NOMBRE"
'    Next i
'
'End Sub
'
'
' ============================================================================
'  APELLIDOS BALEARES (CA-IB)
' ============================================================================

Private Sub CargarApellidosBalear()

'SECCIÓ 1 — Cognoms balears tradicionals

AgregarEntradaDiccionario "CA-IB", "MOLL", "Moll", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MIRALLES", "Miralles", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MUNTANER", "Muntaner", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "SALAS", "Salas", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MASCARÓ", "Mascaro", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "AMENGUAL", "Amengual", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "PONS", "Pons", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MARTORELL", "Martorell", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "VIDAL", "Vidal", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "FERRER", "Ferrer", "APELLIDO", "Llinatge balear tradicional", "Balears"

'AgregarEntradaDiccionario "CA-IB", "MOLL", "Moll", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MIRALLES", "Miralles", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MUNTANER", "Muntaner", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "SALAS", "Salas", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "MASCARÓ", "Mascaro", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "AMENGUAL", "Amengual", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "FERRER", "Ferrer", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "PONS", "Pons", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MARTORELL", "Martorell", "APELLIDO", "Llinatge balear tradicional", "Balears"
'AgregarEntradaDiccionario "CA-IB", "VIDAL", "Vidal", "APELLIDO", "Llinatge balear tradicional", "Balears"


'SECCIÓ 4 — Cognoms exclusivament balears o molt característics

AgregarEntradaDiccionario "CA-IB", "FUSTER", "Fuster", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "MAYOL", "Mayol", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MAYANS", "Mayans", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MOYÀ", "Moya", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MOIÀ", "Moia", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MUNAR", "Munar", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MUNARRO", "Munarro", "APELLIDO", "Llinatge balear", "Mallorca"


'SECCIÓ 2 — Cognoms amb variants balears pròpies

AgregarEntradaDiccionario "CA-IB", "MATEU", "Mateu", "APELLIDO", "Forma balear del llinatge", "Balears"
AgregarEntradaDiccionario "CA-IB", "MARTÍ", "Marti", "APELLIDO", "Forma balear del llinatge", "Balears"
AgregarEntradaDiccionario "CA-IB", "RIGO", "Rigo", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MORA", "Mora", "APELLIDO", "Llinatge balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "MORRO", "Morro", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "MORAGUES", "Moragues", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CANYELLES", "Canyelles", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CANYAMERES", "Canyameres", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CANYELLA", "Canyella", "APELLIDO", "Llinatge balear", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CANYAMAR", "Canyamar", "APELLIDO", "Llinatge balear", "Mallorca"

'AgregarEntradaDiccionario "CA-IB", "MATEU", "Mateu", "APELLIDO", "Forma balear del llinatge", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MARTÍ", "Marti", "APELLIDO", "Forma balear del llinatge", "Balears"
'AgregarEntradaDiccionario "CA-IB", "RIGO", "Rigo", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "MORA", "Mora", "APELLIDO", "Llinatge balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MORRO", "Morro", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "MORAGUES", "Moragues", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "CANYELLES", "Canyelles", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "CANYAMERES", "Canyameres", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "CANYELLA", "Canyella", "APELLIDO", "Llinatge balear", "Mallorca"
'AgregarEntradaDiccionario "CA-IB", "CANYAMAR", "Canyamar", "APELLIDO", "Llinatge balear", "Mallorca"


'SECCIÓ 3 — Cognoms catalans amb fort ús balear

AgregarEntradaDiccionario "CA-IB", "SERRA", "Serra", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "VALLS", "Valls", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "TORRES", "Torres", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "COSTA", "Costa", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "SOLER", "Soler", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ROIG", "Roig", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "PUIG", "Puig", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "FONT", "Font", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "MIR", "Mir", "APELLIDO", "Català comú amb ús balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "GUAL", "Gual", "APELLIDO", "Català comú amb ús balear", "Balears"

'AgregarEntradaDiccionario "CA-IB", "SERRA", "Serra", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "VALLS", "Valls", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "TORRES", "Torres", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "COSTA", "Costa", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "SOLER", "Soler", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "ROIG", "Roig", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "PUIG", "Puig", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "FONT", "Font", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "MIR", "Mir", "APELLIDO", "Català comú amb fort ús balear", "Balears"
'AgregarEntradaDiccionario "CA-IB", "GUAL", "Gual", "APELLIDO", "Català comú amb fort ús balear", "Balears"

'COGNOMS CA-IB — AMPLIACIÓ 1 (Tradicionals i genuïns)

AgregarEntradaDiccionario "CA-IB", "ALCOVER", "Alcover", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ALZINA", "Alzina", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ARBONA", "Arbona", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "ARROM", "Arrom", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BAUZÀ", "Bauza", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BENNÀSSAR", "Bennassar", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BERNAT", "Bernat", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "BINIMELIS", "Binimelis", "APELLIDO", "Llinatge balear antic", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "BORRÀS", "Borras", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "BUSQUETS", "Busquets", "APELLIDO", "Llinatge balear tradicional", "Balears"

'COGNOMS CA-IB — AMPLIACIÓ 2 (Mallorca i Menorca profunds)

AgregarEntradaDiccionario "CA-IB", "CAMPANER", "Campaner", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CAPÓ", "Capo", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CARBONELL", "Carbonell", "APELLIDO", "Llinatge balear tradicional", "Balears"
AgregarEntradaDiccionario "CA-IB", "CATALÀ", "Catala", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CINTAS", "Cintas", "APELLIDO", "Llinatge menorquí tradicional", "Menorca"
AgregarEntradaDiccionario "CA-IB", "CIRER", "Cirer", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "COSTURER", "Costurer", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "COVES", "Coves", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CRESPÍ", "Crespi", "APELLIDO", "Llinatge balear tradicional", "Mallorca"
AgregarEntradaDiccionario "CA-IB", "CUBELLS", "Cubells", "APELLIDO", "Llinatge eivissenc tradicional", "Eivissa"

' COGNOMS CA-IB — AMPLIACIÓ 3 (Eivissa i Formentera)

AgregarEntradaDiccionario "CA-IB", "FERRERET", "Ferreret", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "GUIRADO", "Guirado", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "MARÍ", "Mari", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "RIBAS", "Ribas", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "ROIG", "Roig", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "TORRENT", "Torrent", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "TRUYOL", "Truyol", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "VERICAT", "Vericat", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "VILA", "Vila", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"
AgregarEntradaDiccionario "CA-IB", "VIVÓ", "Vivo", "APELLIDO", "Llinatge pitiús tradicional", "Eivissa"

'COGNOMS CA-IB — AMPLIACIÓ 4 (Catalans amb presència històrica real a les Illes)

AgregarEntradaDiccionario "CA-IB", "ALOMAR", "Alomar", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "BARCELÓ", "Barcelo", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "BENNÀSSAR", "Bennassar", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "CARRIÓ", "Carrio", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "COLL", "Coll", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "ESTADELLA", "Estadella", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "GOMILA", "Gomila", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "JOVER", "Jover", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "MASSANET", "Massanet", "APELLIDO", "Català amb fort arrelament balear", "Balears"
AgregarEntradaDiccionario "CA-IB", "SASTRE", "Sastre", "APELLIDO", "Català amb fort arrelament balear", "Balears"


MsgBox "Fin Apellidos Balear"

End Sub



Private Sub CargarApellidosValencianos()

'SECCIÓ 1 — COGNOMS VALENCIANS TRADICIONALS (NUCLI DUR)

AgregarEntradaDiccionario "CA-VA", "ALBERT", "Albert", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ALFONSO", "Alfonso", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ALMELA", "Almela", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ALMIRALL", "Almirall", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "ANDREU", "Andreu", "APELLIDO", "Llinatge valencià tradicional", "València"

AgregarEntradaDiccionario "CA-VA", "ARAGÓ", "Arago", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "BALAGUER", "Balaguer", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "BENLLOCH", "Benlloch", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "BERENGUER", "Berenguer", "APELLIDO", "Llinatge valencià tradicional", "València"
AgregarEntradaDiccionario "CA-VA", "BLASCO", "Blasco", "APELLIDO", "Llinatge valencià tradicional", "València"


'COGNOMS EXCLUSIVAMENT VALENCIANS — AMPLIACIÓ (MOLT RARS I GENUÏNS)

AgregarEntradaDiccionario "CA-VA", "ADELL", "Adell", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALAPONT", "Alapont", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALBELDA", "Albelda", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALBUIXECH", "Albuixech", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALCÀCER", "Alcacer", "APELLIDO", "Llinatge valencià exclusiu", "València"

AgregarEntradaDiccionario "CA-VA", "ALFARA", "Alfara", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALFONSO", "Alfonso", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALMIRALL", "Almirall", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALMOROIX", "Almoroix", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ALPUENTE", "Alpuente", "APELLIDO", "Llinatge valencià exclusiu", "València"

AgregarEntradaDiccionario "CA-VA", "ARACIL", "Aracil", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ARGENTE", "Argente", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ARRUFAT", "Arrufat", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "ASENSI", "Asensi", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "AUSINA", "Ausina", "APELLIDO", "Llinatge valencià exclusiu", "València"


'COGNOMS EXCLUSIVAMENT VALENCIANS — AMPLIACIÓ 2 (TOPONÍMICS I MEDIEVALS)

AgregarEntradaDiccionario "CA-VA", "BALLESTER", "Ballester", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BARBERÀ", "Barbera", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BARRACHINA", "Barrachina", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BAYARRI", "Bayarri", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BELDA", "Belda", "APELLIDO", "Llinatge valencià antic", "València"

AgregarEntradaDiccionario "CA-VA", "BENAGES", "Benages", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BENET", "Benet", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BENIMELI", "Benimeli", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BERLANGA", "Berlanga", "APELLIDO", "Llinatge valencià antic", "València"
AgregarEntradaDiccionario "CA-VA", "BESALDUCH", "Besalduch", "APELLIDO", "Llinatge valencià antic", "València"


'COGNOMS EXCLUSIVAMENT VALENCIANS — AMPLIACIÓ 3 (L’HORTA, LA RIBERA, LA SAFOR)

AgregarEntradaDiccionario "CA-VA", "BOIRA", "Boira", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BORDA", "Borda", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BORJA", "Borja", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BORT", "Bort", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BRINES", "Brines", "APELLIDO", "Llinatge valencià local", "València"

AgregarEntradaDiccionario "CA-VA", "BROSETA", "Broseta", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BUIGUES", "Buigues", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "BURGUERA", "Burguera", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "CABEDO", "Cabedo", "APELLIDO", "Llinatge valencià local", "València"
AgregarEntradaDiccionario "CA-VA", "CALDUCH", "Calduch", "APELLIDO", "Llinatge valencià local", "València"



'SECCIÓ 2 — COGNOMS EXCLUSIVAMENT VALENCIANS

AgregarEntradaDiccionario "CA-VA", "BORONAT", "Boronat", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "BORRULL", "Borrull", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CALATAYUD", "Calatayud", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CAMPANER", "Campaner", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CARBONELL", "Carbonell", "APELLIDO", "Llinatge valencià exclusiu", "València"

AgregarEntradaDiccionario "CA-VA", "CATALÀ", "Catala", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CERVERA", "Cervera", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CISCAR", "Ciscar", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CLIMENT", "Climent", "APELLIDO", "Llinatge valencià exclusiu", "València"
AgregarEntradaDiccionario "CA-VA", "CORTELL", "Cortell", "APELLIDO", "Llinatge valencià exclusiu", "València"


'SECCIÓ 3 — COGNOMS CATALANS AMB FORT ARRELAMENT VALENCIÀ

AgregarEntradaDiccionario "CA-VA", "FERRER", "Ferrer", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "SERRA", "Serra", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "SOLER", "Soler", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "TORRES", "Torres", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "VILA", "Vila", "APELLIDO", "Català amb ús valencià", "València"

AgregarEntradaDiccionario "CA-VA", "MARTÍ", "Marti", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "MARTORELL", "Martorell", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "PUIG", "Puig", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ROIG", "Roig", "APELLIDO", "Català amb ús valencià", "València"
AgregarEntradaDiccionario "CA-VA", "FONT", "Font", "APELLIDO", "Català amb ús valencià", "València"


'SECCIÓ 4 — COGNOMS VALENCIANS AMB VARIANTS PRÒPIES

AgregarEntradaDiccionario "CA-VA", "MORERA", "Morera", "APELLIDO", "Variant valenciana", "València"
AgregarEntradaDiccionario "CA-VA", "MORANT", "Morant", "APELLIDO", "Variant valenciana", "València"
AgregarEntradaDiccionario "CA-VA", "MORAGUES", "Moragues", "APELLIDO", "Variant valenciana", "València"
'AgregarEntradaDiccionario "CA-VA", "MORERA", "Morera", "APELLIDO", "Variant valenciana", "València"
'AgregarEntradaDiccionario "CA-VA", "MORANT", "Morant", "APELLIDO", "Variant valenciana", "València"

AgregarEntradaDiccionario "CA-VA", "VIDAL", "Vidal", "APELLIDO", "Variant valenciana", "València"
AgregarEntradaDiccionario "CA-VA", "VIDALBERT", "Vidalbert", "APELLIDO", "Variant valenciana", "València"
'AgregarEntradaDiccionario "CA-VA", "VIDALBERT", "Vidalbert", "APELLIDO", "Variant valenciana", "València"
'AgregarEntradaDiccionario "CA-VA", "VIDALBERT", "Vidalbert", "APELLIDO", "Variant valenciana", "València"
'AgregarEntradaDiccionario "CA-VA", "VIDALBERT", "Vidalbert", "APELLIDO", "Variant valenciana", "València"



'COGNOMS VALENCIANS MODERNS D’ÚS REAL — BLOC 1

AgregarEntradaDiccionario "CA-VA", "ALAPONT", "Alapont", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ALBIACH", "Albiach", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ALMELA", "Almela", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "ALMIRALL", "Almirall", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "ALONSO", "Alonso", "APELLIDO", "Ús modern valencià", "València"


'COGNOMS VALENCIANS MODERNS D’ÚS REAL — BLOC 2 (L’Horta, La Ribera, La Safor, La Marina)

AgregarEntradaDiccionario "CA-VA", "BALLESTER", "Ballester", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BARBERÀ", "Barbera", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BARRACHINA", "Barrachina", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BAYARRI", "Bayarri", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BELDA", "Belda", "APELLIDO", "Ús modern valencià", "València"

AgregarEntradaDiccionario "CA-VA", "BENLLOCH", "Benlloch", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BERENGUER", "Berenguer", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BORJA", "Borja", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BORONAT", "Boronat", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "BRINES", "Brines", "APELLIDO", "Ús modern valencià", "València"


'COGNOMS VALENCIANS MODERNS D’ÚS REAL — BLOC 3 (Cognoms molt estesos en l’actualitat)

AgregarEntradaDiccionario "CA-VA", "CABEDO", "Cabedo", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CALDUCH", "Calduch", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CAMPOS", "Campos", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CARRIÓ", "Carrio", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CATALÀ", "Catala", "APELLIDO", "Ús modern valencià", "València"

AgregarEntradaDiccionario "CA-VA", "CISCAR", "Ciscar", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CLIMENT", "Climent", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CORTELL", "Cortell", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CREMADES", "Cremades", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "CUENCA", "Cuenca", "APELLIDO", "Ús modern valencià", "València"


'COGNOMS VALENCIANS MODERNS D’ÚS REAL — BLOC 4 (Cognoms actuals molt freqüents al PV)
'(Els cognoms castellans que són molt valencians en ús real els he inclòs perquè m’has demanat “moderns d’ús real”. Si no els vols, els treus.)

'AgregarEntradaDiccionario "CA-VA", "DOMINGO", "Domingo", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "ESCOBAR", "Escobar", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "FERRANDO", "Ferrando", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "FERRERES", "Ferreres", "APELLIDO", "Ús modern valencià", "València"
AgregarEntradaDiccionario "CA-VA", "FIGUERES", "Figueres", "APELLIDO", "Ús modern valencià", "València"

'AgregarEntradaDiccionario "CA-VA", "GARCIA", "Garcia", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "GIMENO", "Gimeno", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "GÓMEZ", "Gomez", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "GUARDIOLA", "Guardiola", "APELLIDO", "Ús modern valencià", "València"
'AgregarEntradaDiccionario "CA-VA", "IBORRA", "Iborra", "APELLIDO", "Ús modern valencià", "València"

MsgBox "Fin Apellidos Valenciá"
End Sub
