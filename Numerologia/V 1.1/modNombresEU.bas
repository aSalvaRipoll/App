Attribute VB_Name = "modNombresEU"
Option Compare Database
Option Explicit

Public Sub PoblarNombresEU()

    ' ==== NOMBRES VASCOS REALES (1–200 depurados) ====

    ' --- Frecuentes y oficiales ---
    AgregarEntradaDiccionario "EU", "AITOR", "AITOR", "NOMBRE", "Frecuente", "[ai?'to?]"
    AgregarEntradaDiccionario "EU", "IKER", "IKER", "NOMBRE", "Frecuente", "['iker]"
    AgregarEntradaDiccionario "EU", "UNAI", "UNAI", "NOMBRE", "Frecuente", "[u'nai?]"
    AgregarEntradaDiccionario "EU", "ASIER", "ASIER", "NOMBRE", "Frecuente", "[a'sier]"
    AgregarEntradaDiccionario "EU", "JON", "JON", "NOMBRE", "Frecuente", "[jon]"
    AgregarEntradaDiccionario "EU", "MIKEL", "MIKEL", "NOMBRE", "Frecuente", "['mikel]"
    AgregarEntradaDiccionario "EU", "ENEKO", "ENEKO", "NOMBRE", "Frecuente", "[e'neko]"
    AgregarEntradaDiccionario "EU", "GAIZKA", "GAIZKA", "NOMBRE", "Frecuente", "['gais?ka]"
    AgregarEntradaDiccionario "EU", "GORKA", "GORKA", "NOMBRE", "Frecuente", "['gorka]"
    AgregarEntradaDiccionario "EU", "XABIER", "SHABIER", "NOMBRE", "Patrimonial", "[?a'ßie?]"

    AgregarEntradaDiccionario "EU", "ANDER", "ANDER", "NOMBRE", "Frecuente", "['ander]"
    AgregarEntradaDiccionario "EU", "OIER", "OIER", "NOMBRE", "Frecuente", "[o'jer]"
    AgregarEntradaDiccionario "EU", "HAIZEA", "HAIZEA", "NOMBRE", "Frecuente", "[ai?'se.a]"
    AgregarEntradaDiccionario "EU", "AMAIA", "AMAIA", "NOMBRE", "Frecuente", "[a'mai?a]"
    AgregarEntradaDiccionario "EU", "ANE", "ANE", "NOMBRE", "Frecuente", "['ane]"
    AgregarEntradaDiccionario "EU", "NAHIA", "NAHIA", "NOMBRE", "Frecuente", "[na'ia]"
    AgregarEntradaDiccionario "EU", "JUNE", "JUNE", "NOMBRE", "Frecuente", "['june]"
    AgregarEntradaDiccionario "EU", "NEREA", "NEREA", "NOMBRE", "Frecuente", "[ne'?ea]"
    AgregarEntradaDiccionario "EU", "MAIALEN", "MAIALEN", "NOMBRE", "Frecuente", "[mai?a'len]"
    AgregarEntradaDiccionario "EU", "UXUE", "USHUE", "NOMBRE", "Patrimonial", "[u'?ue]"

    AgregarEntradaDiccionario "EU", "ARITZ", "ARITS", "NOMBRE", "Frecuente", "[a'?its?]"
    AgregarEntradaDiccionario "EU", "AIMAR", "AIMAR", "NOMBRE", "Frecuente", "[ai?'ma?]"
    AgregarEntradaDiccionario "EU", "ADUR", "ADUR", "NOMBRE", "Frecuente", "[a'ður]"
    AgregarEntradaDiccionario "EU", "IZARO", "IZARO", "NOMBRE", "Frecuente", "[i'sa?o]"
    AgregarEntradaDiccionario "EU", "IRATI", "IRATI", "NOMBRE", "Frecuente", "[i'?ati]"
    AgregarEntradaDiccionario "EU", "LUR", "LUR", "NOMBRE", "Frecuente", "[lur]"
    AgregarEntradaDiccionario "EU", "LUKEN", "LUKEN", "NOMBRE", "Frecuente", "['luken]"
    AgregarEntradaDiccionario "EU", "LARRAITZ", "LARRAITS", "NOMBRE", "Frecuente", "[la'?ait?s?]"

    ' --- Patrimoniales ---
    AgregarEntradaDiccionario "EU", "ITXASO", "ITCHASO", "NOMBRE", "Patrimonial", "[it??a'so]"
    AgregarEntradaDiccionario "EU", "ITSASNE", "ITSASNE", "NOMBRE", "Patrimonial", "[it?s?as?'ne]"
    AgregarEntradaDiccionario "EU", "ITZIAR", "ITZIAR", "NOMBRE", "Patrimonial", "[it?s?i'a?]"
    AgregarEntradaDiccionario "EU", "ITXARO", "ITCHARO", "NOMBRE", "Patrimonial", "[it??a'?o]"
    AgregarEntradaDiccionario "EU", "XANTI", "SHANTI", "NOMBRE", "Patrimonial", "['?anti]"
    AgregarEntradaDiccionario "EU", "XABINA", "SHABINA", "NOMBRE", "Patrimonial", "[?a'ßina]"
    AgregarEntradaDiccionario "EU", "XARE", "SHARE", "NOMBRE", "Patrimonial", "['?a?e]"
    AgregarEntradaDiccionario "EU", "XARENE", "SHARENE", "NOMBRE", "Patrimonial", "[?a'?ene]"
    AgregarEntradaDiccionario "EU", "XUBAN", "SHUBAN", "NOMBRE", "Patrimonial", "[?u'ßan]"

    ' --- Nombres de naturaleza ---
    AgregarEntradaDiccionario "EU", "OIHANE", "OIHANE", "NOMBRE", "Frecuente", "[oi?'xane]"
    AgregarEntradaDiccionario "EU", "OIHANA", "OIHANA", "NOMBRE", "Frecuente", "[oi?'xana]"
    AgregarEntradaDiccionario "EU", "OIHAN", "OIHAN", "NOMBRE", "Frecuente", "[oi?'xan]"
    AgregarEntradaDiccionario "EU", "OIHANTZ", "OIHANTS", "NOMBRE", "Frecuente", "[oi?'xant?s?]"
    AgregarEntradaDiccionario "EU", "OIHARTZUN", "OIHARTSUN", "NOMBRE", "Frecuente", "[oi?xa?t?s?un]"

    ' --- Otros nombres reales ---
    AgregarEntradaDiccionario "EU", "AMETS", "AMETS", "NOMBRE", "Frecuente", "[a'mets?]"
    AgregarEntradaDiccionario "EU", "AMAIUR", "AMAIUR", "NOMBRE", "Frecuente", "[amai?'u?]"
    AgregarEntradaDiccionario "EU", "ARANTXA", "ARANTXA", "NOMBRE", "Frecuente", "[a'?ant??a]"
    AgregarEntradaDiccionario "EU", "ARAZI", "ARAZI", "NOMBRE", "Frecuente", "[a'?asi]"
    AgregarEntradaDiccionario "EU", "BEÑAT", "BENYAT", "NOMBRE", "Frecuente", "[be'?at]"
    AgregarEntradaDiccionario "EU", "BIDANE", "BIDANE", "NOMBRE", "Frecuente", "[bi'ðane]"
    AgregarEntradaDiccionario "EU", "BIZEN", "BIZEN", "NOMBRE", "Frecuente", "[bi's?en]"
    AgregarEntradaDiccionario "EU", "BIZENTE", "BIZENTE", "NOMBRE", "Frecuente", "[bi's?ente]"
    AgregarEntradaDiccionario "EU", "DANEL", "DANEL", "NOMBRE", "Frecuente", "[da'nel]"
    AgregarEntradaDiccionario "EU", "ELENE", "ELENE", "NOMBRE", "Frecuente", "[e'lene]"
    AgregarEntradaDiccionario "EU", "ELIXABET", "ELIXABET", "NOMBRE", "Frecuente", "[eli'?aßet]"
    AgregarEntradaDiccionario "EU", "ELIXA", "ELIXA", "NOMBRE", "Frecuente", "[e'li?a]"
    AgregarEntradaDiccionario "EU", "ELIXANE", "ELIXANE", "NOMBRE", "Frecuente", "[eli'?ane]"
    AgregarEntradaDiccionario "EU", "ELIXANDRE", "ELIXANDRE", "NOMBRE", "Frecuente", "[eli'?and?e]"
    AgregarEntradaDiccionario "EU", "ELIXANDRO", "ELIXANDRO", "NOMBRE", "Frecuente", "[eli'?and?o]"
    AgregarEntradaDiccionario "EU", "ELIXANDRA", "ELIXANDRA", "NOMBRE", "Frecuente", "[eli'?and?a]"

    ' --- Nombres modernos reales ---
    AgregarEntradaDiccionario "EU", "IRUNE", "IRUNE", "NOMBRE", "Frecuente", "[i'?une]"
    AgregarEntradaDiccionario "EU", "IRUÑA", "IRUNYA", "NOMBRE", "Frecuente", "[i'?u?a]"
    AgregarEntradaDiccionario "EU", "KOLDO", "KOLDO", "NOMBRE", "Frecuente", "['koldo]"
    AgregarEntradaDiccionario "EU", "LARRA", "LARRA", "NOMBRE", "Frecuente", "['lara]"

    ' --- Familia MAITE ---
    AgregarEntradaDiccionario "EU", "MAIDER", "MAIDER", "NOMBRE", "Frecuente", "[mai?'ðe?]"
    AgregarEntradaDiccionario "EU", "MAITANE", "MAITANE", "NOMBRE", "Frecuente", "[mai?'tane]"
    AgregarEntradaDiccionario "EU", "MAITE", "MAITE", "NOMBRE", "Frecuente", "['mai?te]"
    AgregarEntradaDiccionario "EU", "MAITENA", "MAITENA", "NOMBRE", "Frecuente", "[mai?'tena]"
    AgregarEntradaDiccionario "EU", "MAITEN", "MAITEN", "NOMBRE", "Frecuente", "[mai?'ten]"
    
    ' --- Añadir nombres base que faltaban ---

    AgregarEntradaDiccionario "EU", "MIREN", "MIREN", "NOMBRE", "Frecuente", ""      ' Nombre base (forma vasca de María)
    AgregarEntradaDiccionario "EU", "NEKANE", "NEKANE", "NOMBRE", "Frecuente", ""    ' Nombre base, origen religioso/popular

    ' --- Familia ZURI ---
    AgregarEntradaDiccionario "EU", "ZURI", "ZURI", "NOMBRE", "Frecuente", "['su?i]"
    AgregarEntradaDiccionario "EU", "ZURINE", "ZURINE", "NOMBRE", "Frecuente", "[su'?ine]"
    AgregarEntradaDiccionario "EU", "ZURIN", "ZURIN", "NOMBRE", "Frecuente", "[su'?in]"

    ' ==== HIPOCORES VÁLIDOS ====

    AgregarEntradaDiccionario "EU", "MAITETXU", "MAITETXU", "NOMBRE", "Hipocorístico", ""     ' De Maite
    AgregarEntradaDiccionario "EU", "MIRENTXU", "MIRENTXU", "NOMBRE", "Hipocorístico", ""     ' De Miren
    AgregarEntradaDiccionario "EU", "ANETXU", "ANETXU", "NOMBRE", "Hipocorístico", ""         ' De Ane
    AgregarEntradaDiccionario "EU", "NEKETXU", "NEKETXU", "NOMBRE", "Hipocorístico", ""       ' De Nekane

    AgregarEntradaDiccionario "EU", "AITORTXU", "AITORTXU", "NOMBRE", "Hipocorístico", ""     ' De Aitor
    AgregarEntradaDiccionario "EU", "JONETXU", "JONETXU", "NOMBRE", "Hipocorístico", ""       ' De Jon
    AgregarEntradaDiccionario "EU", "MIKELTXU", "MIKELTXU", "NOMBRE", "Hipocorístico", ""     ' De Mikel

Debug.Print "Fin"

End Sub


'Public Sub PoblarApellidosEU()
'
'' ==== BLOQUE 1 — APELLIDOS 1–50 ====
'
'AgregarEntradaDiccionario "EU", "ETXEBERRIA", "ECHEBERRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEA", "GOICOECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIZARRETA", "ZUBIZARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URRUTIA", "URRUTIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AGIRRE", "AGIRRE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARRIETA", "ARRIETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "LARRAÑAGA", "LARRAÑAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "MENDIZABAL", "MENDIZABAL", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEBESTE", "ECHEBESTE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXENIKE", "ECHENIKE", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "ETXEPARE", "ECHEPARE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEGARAI", "ECHEGARAI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXARRI", "ECHARRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXANIZ", "ECHANIZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEBARRIA", "ECHEBARRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEZARRETA", "ECHEZARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEZAR", "ECHEZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEBERRI", "ECHEBERRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEBERRIAZAR", "ECHEBERRIAZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ETXEBERRIETXEA", "ECHEBERRIETCHEA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "ZABALA", "ZABALA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZALDUA", "ZALDUA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIA", "ZUBIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIRI", "ZUBIRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIAURRE", "ZUBIAURRE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIMENDI", "ZUBIMENDI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIMURU", "ZUBIMURU", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIZAR", "ZUBIZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIZAGA", "ZUBIZAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZUBIZARRETA", "ZUBIZARRETA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "ZARAUTZ", "ZARAUZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARATE", "ZARATE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARANDONA", "ZARANDONA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARANDIETA", "ZARANDIETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARATEGUI", "ZARATEGUI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARATIEGUI", "ZARATIEGUI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARATIEZ", "ZARATIEZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARAUZA", "ZARAUZA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARAUZAGA", "ZARAUZAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ZARAUZALDE", "ZARAUZALDE", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "XABIER", "SHABIER", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XIMUN", "SHIMUN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XENPELAR", "SHENPELAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XARETA", "SHARETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XUBIRI", "SHUBIRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XURRUT", "SHURRUT", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XURRUKA", "SHURRUKA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XURMENDI", "SHURMENDI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XURIO", "SHURIO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "XABALO", "SHABALO", "APELLIDO", "Vasco", ""
'
'' ==== BLOQUE 2 — APELLIDOS 51–100 ====
'
'AgregarEntradaDiccionario "EU", "URQUIZU", "URKIZU", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIAGA", "URKIAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIOLA", "URKIOLA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIBARRI", "URKIBARRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZAGA", "URKIZAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZAR", "URKIZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZALDE", "URKIZALDE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZARRA", "URKIZARRA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZAL", "URKIZAL", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URQUIZAGA", "URKIZAGA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "ARAMBURU", "ARAMBURU", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARAMBERRI", "ARAMBERRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANBURU", "ARANBURU", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZADI", "ARANZADI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZAMENDI", "ARANZAMENDI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZATEGUI", "ARANZATEGUI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZIBAR", "ARANZIBAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZUBIA", "ARANZUBIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZUGA", "ARANZUGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ARANZUGALDE", "ARANZUGALDE", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "IBARROLA", "IBARROLA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRA", "IBARRA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDO", "IBARRONDO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRURI", "IBARRURI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOA", "IBARRONDOA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOZ", "IBARRONDOZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOAGA", "IBARRONDOAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOETXEA", "IBARRONDOECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOZAR", "IBARRONDOZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IBARRONDOZAGA", "IBARRONDOZAGA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "OTEGI", "OTEGI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOA", "OCHOA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXANDIANO", "OCHANDIANO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARAN", "OCHARAN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARTE", "OCHARTE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARRETA", "OCHARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEBERRIA", "OCHOAECHEBERRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEA", "OCHOAECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEBARRIA", "OCHOAECHEBARRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEGARAI", "OCHOAECHEGARAI", "APELLIDO", "Vasco", ""
'
'' ==== BLOQUE 3 — APELLIDOS 101–150 ====
'
'AgregarEntradaDiccionario "EU", "AIZPURUA", "AIZPURUA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZPEOLEA", "AIZPEOLEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORBE", "AIZKORBE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRETA", "AIZKORRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRI", "AIZKORRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRAGA", "AIZKORRAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRIZAR", "AIZKORRIZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRIZAGA", "AIZKORRIZAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRIALDE", "AIZKORRIALDE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AIZKORRIBAR", "AIZKORRIBAR", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "IRIGOIEN", "IRIGOIEN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARAY", "IRIGARAY", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARATE", "IRIGARATE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZABAL", "IRIGARZABAL", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZAGA", "IRIGARZAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZALDE", "IRIGARZALDE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZARRA", "IRIGARZARRA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZETA", "IRIGARZETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZIBAR", "IRIGARZIBAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "IRIGARZUBIA", "IRIGARZUBIA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "OTEGI", "OTEGI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOA", "OCHOA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXANDIANO", "OCHANDIANO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARAN", "OCHARAN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARTE", "OCHARTE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXARRETA", "OCHARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEBERRIA", "OCHOAECHEBERRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEA", "OCHOAECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEBARRIA", "OCHOAECHEBARRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "OTXOAETXEGARAI", "OCHOAECHEGARAI", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "GOENAGA", "GOENAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEBARRIA", "GOICOECHEBARRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEBERRI", "GOICOECHEBERRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEBERRIA", "GOICOECHEBERRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEGARAI", "GOICOECHEGARAI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEZAR", "GOICOECHEZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEZARRETA", "GOICOECHEZARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEZUBI", "GOICOECHEZUBI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEZUBIA", "GOICOECHEZUBIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIKOETXEZUBIALDE", "GOICOECHEZUBIALDE", "APELLIDO", "Vasco", ""
'
'' ==== BLOQUE 4 — APELLIDOS 151–200 ====
'
'AgregarEntradaDiccionario "EU", "ELOSEGUI", "ELOSEGUI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEBARRI", "ELOSEBARRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZAR", "ELOSEZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZARRETA", "ELOSEZARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZARRA", "ELOSEZARRA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZUBI", "ELOSEZUBI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZUBIA", "ELOSEZUBIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZUBIZAR", "ELOSEZUBIZAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZUBIZARRETA", "ELOSEZUBIZARRETA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ELOSEZUBIALDE", "ELOSEZUBIALDE", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "BENGOETXEA", "BENGOECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "AURREKOETXEA", "AURRECOECHEA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUIZABAL", "EGUIZABAL", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUIA", "EGUIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUIARTE", "EGUIARTE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUIBAR", "EGUIBAR", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUIN", "EGUIN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUINOA", "EGUINOA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUINOZ", "EGUINOZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "EGUREN", "EGUREN", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "GARAI", "GARAI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GARATE", "GARATE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GARITANO", "GARITANO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GAZTELU", "GAZTELU", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GAZTANAGA", "GAZTANAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GAZTAÑAGA", "GAZTAÑAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOENAGA", "GOENAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIRI", "GOIRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIRIA", "GOIRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "GOIRIZELAIA", "GOIRIZELAIA", "APELLIDO", "Vasco", ""
'
'AgregarEntradaDiccionario "EU", "ITURRI", "ITURRI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ITURRIA", "ITURRIA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ITURRIZ", "ITURRIZ", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ITURMENDI", "ITURMENDI", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "ITURRALDE", "ITURRALDE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "LIZARRALDE", "LIZARRALDE", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "LIZARRAGA", "LIZARRAGA", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "LIZASO", "LIZASO", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "LIZASOAIN", "LIZASOAIN", "APELLIDO", "Vasco", ""
'AgregarEntradaDiccionario "EU", "URDANGARIN", "URDANGARIN", "APELLIDO", "Vasco", ""
'
'Debug.Print "Fin"
'End Sub

Sub PoblarApellidosVascos()

    ' === APELLIDOS VASCOS NATIVOS (FONÉTICA ASCII-SAFE) ===
    ' Reglas aplicadas:
    ' TX ? CH
    ' TZ ? Z
    ' TS ? S
    ' X ? SH
    ' RR ? RR
    ' Z ? Z

    ' ============================================================
    ' ===============   BLOQUE 1 — APELLIDOS 1–50   ===============
    ' ============================================================

    AgregarEntradaDiccionario "EU", "ETXEBERRIA", "ECHEBERRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GOIKOETXEA", "GOICOECHEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIZARRETA", "ZUBIZARRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URRUTIA", "URRUTIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AGIRRE", "AGIRRE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARRIETA", "ARRIETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "LARRAÑAGA", "LARRAÑAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "MENDIZABAL", "MENDIZABAL", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEBESTE", "ECHEBESTE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXENIKE", "ECHENIKE", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "ETXEPARE", "ECHEPARE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEGARAI", "ECHEGARAI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXARRI", "ECHARRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXANIZ", "ECHANIZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEBARRIA", "ECHEBARRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEZARRETA", "ECHEZARRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEZAR", "ECHEZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEBERRI", "ECHEBERRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEBERRIAZAR", "ECHEBERRIAZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ETXEBERRIETXEA", "ECHEBERRIETCHEA", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "ZABALA", "ZABALA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZALDUA", "ZALDUA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIA", "ZUBIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIRI", "ZUBIRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIAURRE", "ZUBIAURRE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIMENDI", "ZUBIMENDI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIMURU", "ZUBIMURU", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIZAR", "ZUBIZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIZAGA", "ZUBIZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZUBIZARRETA", "ZUBIZARRETA", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "ZARAUTZ", "ZARAUZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARATE", "ZARATE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARANDONA", "ZARANDONA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARANDIETA", "ZARANDIETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARATEGUI", "ZARATEGUI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARATIEGUI", "ZARATIEGUI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARATIEZ", "ZARATIEZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARAUZA", "ZARAUZA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARAUZAGA", "ZARAUZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ZARAUZALDE", "ZARAUZALDE", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "XABIER", "SHABIER", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XIMUN", "SHIMUN", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XENPELAR", "SHENPELAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XARETA", "SHARETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XUBIRI", "SHUBIRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XURRUT", "SHURRUT", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XURRUKA", "SHURRUKA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XURMENDI", "SHURMENDI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XURIO", "SHURIO", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "XABALO", "SHABALO", "APELLIDO", "Vasco", ""


    ' ============================================================
    ' ===============   BLOQUE 2 — APELLIDOS 51–100   ===============
    ' ============================================================

    AgregarEntradaDiccionario "EU", "URQUIZU", "URKIZU", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIAGA", "URKIAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIOLA", "URKIOLA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIBARRI", "URKIBARRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZAGA", "URKIZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZAR", "URKIZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZALDE", "URKIZALDE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZARRA", "URKIZARRA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZAGA", "URKIZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URQUIZAL", "URKIZAL", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "ARAMBURU", "ARAMBURU", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARAMBERRI", "ARAMBERRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANBURU", "ARANBURU", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZADI", "ARANZADI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZAMENDI", "ARANZAMENDI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZATEGUI", "ARANZATEGUI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZIBAR", "ARANZIBAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZUBIA", "ARANZUBIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZUGA", "ARANZUGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ARANZUGALDE", "ARANZUGALDE", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "IBARROLA", "IBARROLA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRA", "IBARRA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDO", "IBARRONDO", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRURI", "IBARRURI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOA", "IBARRONDOA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOZ", "IBARRONDOZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOAGA", "IBARRONDOAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOETXEA", "IBARRONDOECHEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOZAR", "IBARRONDOZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IBARRONDOZAGA", "IBARRONDOZAGA", "APELLIDO", "Vasco", ""


    ' ============================================================
    ' ===============   BLOQUE 3 — APELLIDOS 101–150   ===============
    ' ============================================================

    AgregarEntradaDiccionario "EU", "OTEGI", "OTEGI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXOA", "OCHOA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXANDIANO", "OCHANDIANO", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXARAN", "OCHARAN", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXARTE", "OCHARTE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXARRETA", "OCHARRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXOAETXEBERRIA", "OCHOAECHEBERRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXOAETXEA", "OCHOAECHEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXOAETXEBARRIA", "OCHOAECHEBARRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "OTXOAETXEGARAI", "OCHOAECHEGARAI", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "AIZPURUA", "AIZPURUA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZPEOLEA", "AIZPEOLEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORBE", "AIZKORBE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRETA", "AIZKORRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRI", "AIZKORRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRAGA", "AIZKORRAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRIZAR", "AIZKORRIZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRIZAGA", "AIZKORRIZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRIALDE", "AIZKORRIALDE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AIZKORRIBAR", "AIZKORRIBAR", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "IRIGOIEN", "IRIGOIEN", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARAY", "IRIGARAY", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARATE", "IRIGARATE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZABAL", "IRIGARZABAL", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZAGA", "IRIGARZAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZALDE", "IRIGARZALDE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZARRA", "IRIGARZARRA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZETA", "IRIGARZETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZIBAR", "IRIGARZIBAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "IRIGARZUBIA", "IRIGARZUBIA", "APELLIDO", "Vasco", ""


    ' ============================================================
    ' ===============   BLOQUE 4 — APELLIDOS 151–200   ===============
    ' ============================================================

    AgregarEntradaDiccionario "EU", "ELOSEGUI", "ELOSEGUI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEBARRI", "ELOSEBARRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZAR", "ELOSEZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZARRETA", "ELOSEZARRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZARRA", "ELOSEZARRA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZUBI", "ELOSEZUBI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZUBIA", "ELOSEZUBIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZUBIZAR", "ELOSEZUBIZAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZUBIZARRETA", "ELOSEZUBIZARRETA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ELOSEZUBIALDE", "ELOSEZUBIALDE", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "BENGOETXEA", "BENGOECHEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "AURREKOETXEA", "AURRECOECHEA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUIZABAL", "EGUIZABAL", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUIA", "EGUIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUIARTE", "EGUIARTE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUIBAR", "EGUIBAR", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUIN", "EGUIN", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUINOA", "EGUINOA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUINOZ", "EGUINOZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "EGUREN", "EGUREN", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "GARAI", "GARAI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GARATE", "GARATE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GARITANO", "GARITANO", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GAZTelu", "GAZTELU", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GAZTANAGA", "GAZTANAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GAZTAÑAGA", "GAZTAÑAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GOENAGA", "GOENAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GOIRI", "GOIRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GOIRIA", "GOIRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "GOIRIZELAIA", "GOIRIZELAIA", "APELLIDO", "Vasco", ""

    AgregarEntradaDiccionario "EU", "ITURRI", "ITURRI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ITURRIA", "ITURRIA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ITURRIZ", "ITURRIZ", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ITURMENDI", "ITURMENDI", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "ITURRALDE", "ITURRALDE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "LIZARRALDE", "LIZARRALDE", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "LIZARRAGA", "LIZARRAGA", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "LIZASO", "LIZASO", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "LIZASOAIN", "LIZASOAIN", "APELLIDO", "Vasco", ""
    AgregarEntradaDiccionario "EU", "URDANGARIN", "URDANGARIN", "APELLIDO", "Vasco", ""

Debug.Print "Fin"

End Sub

