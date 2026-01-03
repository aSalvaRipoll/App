VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmEntradaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Public IdiomaNom As clsIdioma
Public IdiomaApe1 As clsIdioma
Public IdiomaApe2 As clsIdioma


Dim fonNom As String
Dim FonApe1 As String
Dim FonApe2 As String

Private Sub cboGenero_Click()

'SELECT
'    G.ID_Genero,
'    Nz(I.Icono,'') & Nz(I2.Icono,'') AS IconoGenero,
'    G.Genero
'FROM
'    tbmGeneros As G
'    LEFT JOIN tbmIconos AS I ON I.Codigo = G.CodigoIconoBase
'    LEFT JOIN tbmIconos AS I2 ON I2.Codigo = G.CodigoIconoExtra
'WHERE
'    G.Activo = True
'ORDER BY
'    G.Genero;

End Sub

Private Sub cmdBorrar_Click()
    BorraForm
End Sub

Private Sub cmdBuscar_Click()

    BorraForm

    DoCmd.OpenForm "frmSelectorPersonas", acNormal
    
End Sub

Private Sub cmdCalcular_Click()

    Call CargarClaseDesdeFormulario(Me)
    Call CalcularNumerologia(Persona)
    Call VolcarResultadosAlFormulario

End Sub

Private Sub cmdConvertir_Click()

    If Len(Trim$(Me.txtNombreBase)) > 0 Then
        fonNom = MF_Convertir(Me.txtNombreBase, IdiomaNom, False)
    End If

    If Len(Trim$(Me.txtApeBase1)) > 0 Then
        FonApe1 = MF_Convertir(Me.txtApeBase1, IdiomaApe1, True)
    End If

    If Len(Trim$(Me.txtApeBase2)) > 0 Then
        FonApe2 = MF_Convertir(Me.txtApeBase2, IdiomaApe2, True)
    End If

    Dim resultado As String
    resultado = fonNom

    If FonApe1 <> "" Then resultado = resultado & " " & FonApe1
    If FonApe2 <> "" Then resultado = resultado & " " & FonApe2

    Me.txtFonetico = resultado

    
    
    
    ' 2. Guardar en la tabla
    GuardarFoneticaSmart


    ' 1. Pasar datos del formulario a la clase
    Call CargarClaseDesdeFormulario(Me)

    ' 2. Guardar en la tabla
    Call GuardarPersonaEnTabla

    ' 3. Refrescar el formulario (opcional)
    Call RefrescarDesdeClase


'    On Error GoTo ErrHandler
'    Exit Sub
'
'ErrHandler:
'    MsgBox "Error en la conversión fonética: " & Err.Description, vbExclamation, "Conversor Fonético"
End Sub

Private Sub cmdGuardar_Click()

    ' 1. Pasar datos del formulario a la clase
    Call CargarClaseDesdeFormulario(Me)

    ' 2. Guardar en la tabla
    Call GuardarPersonaEnTabla

    ' 3. Refrescar el formulario (opcional)
    Call RefrescarDesdeClase

    MsgBox "Datos guardados correctamente.", vbInformation

End Sub

Private Sub cmdIdiomaApe1_Click()

    Dim nuevo As clsIdioma
    
    Set IdiomaActual = IdiomaApe1
    CampoDestino = "Ape1"
    
    'Set nuevo = SeleccionarIdioma("Ape1", IdiomaApe1)
    Set nuevo = SeleccionarIdioma(IdiomaApe1)

    If Not nuevo Is Nothing Then
        Set IdiomaApe1 = nuevo
        Call CambiarIcono(Me.cmdIdiomaApe1, IdiomaApe1.Abreviado)
    End If

End Sub

Private Sub cmdIdiomaApe2_Click()

    Dim nuevo As clsIdioma
    
    Set IdiomaActual = IdiomaApe2
    CampoDestino = "Ape2"
    
    'Set nuevo = SeleccionarIdioma("Ape2", IdiomaApe2)
    Set nuevo = SeleccionarIdioma(IdiomaApe2)

    If Not nuevo Is Nothing Then
        Set IdiomaApe2 = nuevo
        Call CambiarIcono(Me.cmdIdiomaApe2, IdiomaApe2.Abreviado)
    End If

End Sub

Private Sub cmdIdiomaNom_Click()
    
    Dim nuevo As clsIdioma
    
    Set IdiomaActual = IdiomaNom
    CampoDestino = "Nom"
    
    'Set nuevo = SeleccionarIdioma("Nom", IdiomaNom)
    Set nuevo = SeleccionarIdioma(IdiomaNom)

    If Not nuevo Is Nothing Then
        Set IdiomaNom = nuevo
        Call CambiarIcono(Me.cmdIdiomaNom, IdiomaNom.Abreviado)
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
    
    On Error GoTo Finally
    
    Me.Painting = False
    Me.fraSistema = 2
    ActualizarOpcionesSistema

Finally:
    Me.Painting = True
    
End Sub

Private Sub Form_Load()
    
    If colIdiomas Is Nothing Then
        Call CargarIdiomas
    End If
    
    
    BorraForm
    
'    ' Inicializar idiomas por defecto si están vacíos
'    If IdiomaNom Is Nothing Then
'        Set IdiomaNom = colIdiomas("0")   ' Español, por ejemplo
'        Call CambiarIcono(Me.cmdIdiomaNom, IdiomaNom.Abreviado)
'    End If
'
'    If IdiomaApe1 Is Nothing Then
'        Set IdiomaApe1 = colIdiomas("0")
'        Call CambiarIcono(Me.cmdIdiomaApe1, IdiomaApe1.Abreviado)
'    End If
'
'    If IdiomaApe2 Is Nothing Then
'        Set IdiomaApe2 = colIdiomas("0")
'        Call CambiarIcono(Me.cmdIdiomaApe2, IdiomaApe2.Abreviado)
'    End If
'
'    ' Inicializar Sistema de análisis
'    Me.fraSistema = 2
'
'    ' Inicializar Tarot
'    Me.fraTarot = 2
'
'    '--- Limpiar campos de salida ---
'    Me.txtFonetico = ""
'    Me.txtNormal = ""
'
'    '--- Limpiar campos de entrada ---
'    Me.txtNombreBase = ""
'    Me.txtApeBase1 = ""
'    Me.txtApeBase2 = ""
'    Me.txtFechaNac = Null
'    Me.cboGenero = Null
'
'    '--- Activar botones de idioma ---
'    Me.cmdIdiomaNom.Enabled = True
'    Me.cmdIdiomaApe1.Enabled = True
'    Me.cmdIdiomaApe2.Enabled = True
'
'    '--- Colocar foco inicial ---
'    Me.txtNombreBase.SetFocus

    Set Persona = New clsPersona

End Sub

Private Sub fraSistema_Click()

    ActualizarOpcionesSistema

End Sub

Private Sub ActualizarOpcionesSistema()
    
    With Me
        Dim activar As Boolean
        activar = (.fraSistema = 2)
        
        .chkHmuda.Enabled = activar
        .chkHmuda.Value = activar
        
        .chkUmuda.Enabled = activar
        .chkUmuda.Value = activar
    End With

End Sub
    

Public Sub CambiarIcono(btn As CommandButton, Abrev As String)
    Dim ruta As String
    
    If Abrev & "" = "" Then
        ruta = CurrentProject.Path & "\img\Test.png"
    Else
    'ruta = CurrentProject.Path & "\img\" & Abrev & ".png"
        ruta = CurrentProject.Path & "\img\" & Abrev & ".ico"
    End If
    
    If Dir(ruta) = "" Then
        'ruta = CurrentProject.Path & "\img\other.png"
        ruta = CurrentProject.Path & "\img\Otros.ico"
    End If

    btn.Picture = ruta
End Sub
Private Function SeleccionarIdioma(IdiomaActual As clsIdioma) As clsIdioma

'Private Function SeleccionarIdioma(campo As String, IdiomaActual As clsIdioma) As clsIdioma

    DoCmd.OpenForm "frmIdioma", , , , , acDialog 'acHidden

    ' Pasar el idioma ANTES de que el usuario interactúe
    'Set Form_frmIdioma.IdiomaActual = IdiomaActual
    'Form_frmIdioma.CampoDestino = campo
    
    ' Mostrar modal
    
    'Forms!frmIdioma.Modal = True
    'Forms!frmIdioma.Visible = True
    
    ' Recoger resultado
    'If Forms!frmIdioma.Tag <> "" Then
    If Not IdiomaSeleccionado Is Nothing Then
        'Set SeleccionarIdioma = Forms!frmIdioma.IdiomaSeleccionado
        Set SeleccionarIdioma = IdiomaSeleccionado
    Else
        Set SeleccionarIdioma = IdiomaActual
    End If

End Function

Private Sub BorraForm()
        
    '--- Limpiar campos de entrada ---
    Me.txtNombreBase = ""
    Me.txtApeBase1 = ""
    Me.txtApeBase2 = ""
    Me.txtFechaNac = Null
    Me.cboGenero = Null
        
    '--- Activar botones de idioma ---
    Me.cmdIdiomaNom.Enabled = True
    Me.cmdIdiomaApe1.Enabled = True
    Me.cmdIdiomaApe2.Enabled = True
        
        
    '--- Inicializar botones idioma ---
    Set IdiomaNom = colIdiomas("0")   ' Español, por ejemplo
    Call CambiarIcono(Me.cmdIdiomaNom, IdiomaNom.Abreviado)

    Set IdiomaApe1 = colIdiomas("0")
    Call CambiarIcono(Me.cmdIdiomaApe1, IdiomaApe1.Abreviado)

    Set IdiomaApe2 = colIdiomas("0")
    Call CambiarIcono(Me.cmdIdiomaApe2, IdiomaApe2.Abreviado)
    
    '--- Inicializar Sistema de análisis ---
    Me.fraSistema = 2
    
    '--- Inicializar Tarot ---
    Me.fraTarot = 2
    
    
    '--- Limpiar campos de salida ---
    Me.txtFonetico = ""
    Me.txtNormal = ""
    
    
    '--- Colocar foco inicial ---
    Me.txtNombreBase.SetFocus

    Set Persona = New clsPersona

End Sub


Public Sub RefrescarDesdeClase()

    ' ============================================================
    '   ZONA 1 — DATOS PERSONALES
    ' ============================================================

    Me.txtIDPersona = Persona.ID_Persona
    Me.txtNombreBase = Persona.Nombre
    Me.txtApeBase1 = Persona.Ape1
    Me.txtApeBase2 = Persona.Ape2

    If Persona.FechaNacimiento = 0 Then
        Me.txtFechaNac = Null
    Else
        Me.txtFechaNac = Persona.FechaNacimiento
    End If

    ' Combo de género (usa ID_Genero)
    Me.cboGenero = Persona.ID_Genero

    ' Idiomas
'    Me.cboIdiomaNombre = Persona.IdiomaNombre
'    Me.cboIdiomaApe1 = Persona.IdiomaApe1
'    Me.cboIdiomaApe2 = Persona.IdiomaApe2


    ' --- Idioma del Nombre ---
    If Not Persona.IdiomaNombre Is Nothing Then
        Set IdiomaNom = Persona.IdiomaNombre
    Else
        Set IdiomaNom = colIdiomas("0")
    End If
    Call CambiarIcono(Me.cmdIdiomaNom, IdiomaNom.Abreviado)
    
    ' --- Idioma del Apellido 1 ---
    If Not Persona.IdiomaApe1 Is Nothing Then
        Set IdiomaApe1 = Persona.IdiomaApe1
    Else
        Set IdiomaApe1 = colIdiomas("0")
    End If
    Call CambiarIcono(Me.cmdIdiomaApe1, IdiomaApe1.Abreviado)
        
    ' --- Idioma del Apellido 2 ---
    If Not Persona.IdiomaApe2 Is Nothing Then
        Set IdiomaApe2 = Persona.IdiomaApe2
    Else
        Set IdiomaApe1 = colIdiomas("0")
    End If
    Call CambiarIcono(Me.cmdIdiomaApe2, IdiomaApe2.Abreviado)
    
    
    ' ============================================================
    '   ZONA 2 — FONÉTICA
    ' ============================================================

    fonNom = Persona.FonNombre
    FonApe1 = Persona.FonApe1
    FonApe2 = Persona.FonApe2

    ' Resultado fonético combinado
    Me.txtFonetico = Trim$(fonNom & " " & FonApe1 & " " & FonApe2)


    ' ============================================================
    '   ZONA 3 — PARÁMETROS DE ANÁLISIS
    ' ============================================================

    Me.fraSistema = Persona.Sistema
    Me.fraTarot = Persona.Tarot
    Me.chkHmuda = Persona.UsarHmuda
    Me.chkUmuda = Persona.UsarUmuda


    ' ============================================================
    '   GESTIÓN (opcional mostrar)
    ' ============================================================

'    If Me.Controls("txtFechaAlta") Is Nothing = False Then
'        If Persona.FechaAlta = 0 Then
'            Me.txtFechaAlta = Null
'        Else
'            Me.txtFechaAlta = Persona.FechaAlta
'        End If
'    End If
'
'    If Me.Controls("txtFechaModificacion") Is Nothing = False Then
'        If Persona.FechaModificacion = 0 Then
'            Me.txtFechaModificacion = Null
'        Else
'            Me.txtFechaModificacion = Persona.FechaModificacion
'        End If
'    End If

End Sub


Public Sub CargarClaseDesdeFormulario(frm As Form)

    ' ============================================================
    '   ZONA 1 — DATOS PERSONALES
    ' ============================================================

    Persona.ID_Persona = Nz(frm.txtIDPersona, 0)
    Persona.Nombre = Nz(frm.txtNombreBase, "")
    Persona.Ape1 = Nz(frm.txtApeBase1, "")
    Persona.Ape2 = Nz(frm.txtApeBase2, "")

    If IsNull(frm.txtFechaNac) Or frm.txtFechaNac = "" Then
        Persona.FechaNacimiento = 0
    Else
        Persona.FechaNacimiento = frm.txtFechaNac
    End If

    Persona.ID_Genero = Nz(frm.cboGenero, 0)


    ' ============================================================
    '   ZONA 2 — IDIOMAS (OBJETOS clsIdioma)
    ' ============================================================

    ' Idioma del Nombre
    If Not frm.IdiomaNom Is Nothing Then
        Set Persona.IdiomaNombre = frm.IdiomaNom
    Else
        Set Persona.IdiomaNombre = Nothing
    End If

    ' Idioma del Apellido 1
    If Not frm.IdiomaApe1 Is Nothing Then
        Set Persona.IdiomaApe1 = frm.IdiomaApe1
    Else
        Set Persona.IdiomaApe1 = Nothing
    End If

    ' Idioma del Apellido 2
    If Not frm.IdiomaApe2 Is Nothing Then
        Set Persona.IdiomaApe2 = frm.IdiomaApe2
    Else
        Set Persona.IdiomaApe2 = Nothing
    End If


'    ' ============================================================
'    '   ZONA 3 — FONÉTICA
'    ' ============================================================
'
'    Persona.FonNombre = Nz(frm.txtFonNombre, "")
'    Persona.FonApe1 = Nz(frm.txtFonApe1, "")
'    Persona.FonApe2 = Nz(frm.txtFonApe2, "")

    ' ============================================================
    '   ZONA 3 — FONÉTICA (solo variables internas)
    ' ============================================================
    
    Persona.FonNombre = fonNom
    Persona.FonApe1 = FonApe1
    Persona.FonApe2 = FonApe2


    ' ============================================================
    '   ZONA 4 — PARÁMETROS DE ANÁLISIS
    ' ============================================================

    Persona.Sistema = frm.fraSistema
    Persona.Tarot = frm.fraTarot
    Persona.UsarHmuda = frm.chkHmuda
    Persona.UsarUmuda = frm.chkUmuda


    ' ============================================================
    '   ZONA 5 — GESTIÓN
    ' ============================================================

    'If Not IsNull(frm.txtFechaAlta) Then
    '    Persona.FechaAlta = frm.txtFechaAlta
    'End If

    ' FechaModificacion se asigna al guardar
    ' No se toca aquí

End Sub

