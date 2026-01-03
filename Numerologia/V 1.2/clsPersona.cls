VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' ============================================================
'   CLASE PERSONA — ESTRUCTURA BASE
'   (Datos personales, 5 números, adicionales, herencias,
'    ciclos, pináculos y desafíos)
' ============================================================

Option Compare Database
Option Explicit

' ============================================================
'   1. DATOS PERSONALES
' ============================================================
Private mNombre As String
Private mApe1 As String
Private mApe2 As String
Private mFechaNac As Date

' Fonética por palabra
Private mFonNombre As SalidaDatos
Private mFonApe1 As SalidaDatos
Private mFonApe2 As SalidaDatos

' Valores fonéticos
Private mVocalesTotal As Integer
Private mConsonantesTotal As Integer
Private mVocalesApe1 As Integer
Private mVocalesApe2 As Integer
Private mConsonantesApe1 As Integer
Private mConsonantesApe2 As Integer


' ============================================================
'   2. NÚMEROS PRINCIPALES (tResultado)
' ============================================================
Private mAlma As tResultado
Private mPersonalidad As tResultado
Private mDestino As tResultado
Private mCamino As tResultado
Private mMadurez As tResultado


' ============================================================
'   3. NÚMEROS ADICIONALES
' ============================================================
Private mDiaNacimiento As tResultado
Private mAnoPersonal As tResultado
Private mEdadPersonal As tResultado
Private mPlanoExpresion As Byte
Private mPrimeraLetra As tResultado
Private mPrimeraVocal As tResultado
Private mPrimeraConsonante As tResultado
Private mPoder As tResultado
Private mRespuestaSubconsciente As Byte


' ============================================================
'   4. HERENCIAS
' ============================================================
Private mHerenciaPaterna As tResultado
Private mHerenciaMaterna As tResultado
Private mHerenciaFinal As tResultado
Private mTipoHerencia As Byte        '1 o 3
Private mNotaHerencia As String


' ============================================================
'   5. CICLOS (4 ciclos)
' ============================================================
Private mCiclo(1 To 4) As tResultado
Private mCicloInicio(1 To 4) As Date
Private mCicloFin(1 To 4) As Date
Private mCicloActual As Byte


' ============================================================
'   6. PINÁCULOS (4)
' ============================================================
Private mPinaculo(1 To 4) As tResultado
Private mPinaculoInicio(1 To 4) As Date
Private mPinaculoFin(1 To 4) As Date
Private mPinaculoActual As Byte


' ============================================================
'   7. DESAFÍOS (4)
' ============================================================
Private mDesafio(1 To 4) As tResultado
Private mDesafioActual As Byte



' ============================================================
'   PROPIEDADES PÚBLICAS (solo lectura)
' ============================================================

Public Property Get Nombre() As String
    Nombre = mNombre
End Property
Public Property Get Ape1() As String
    Ape1 = mApe1
End Property
Public Property Get Ape2() As String
    Ape2 = mApe2
End Property

Public Property Get FechaNacimiento() As Date
    FechaNacimiento = mFechaNac
End Property

Public Property Get Alma() As tResultado: Alma = mAlma: End Property
Public Property Get Personalidad() As tResultado: Personalidad = mPersonalidad: End Property
Public Property Get Destino() As tResultado: Destino = mDestino: End Property
Public Property Get Camino() As tResultado: Camino = mCamino: End Property
Public Property Get Madurez() As tResultado: Madurez = mMadurez: End Property

Public Property Get HerenciaPaterna() As tResultado: HerenciaPaterna = mHerenciaPaterna: End Property
Public Property Get HerenciaMaterna() As tResultado: HerenciaMaterna = mHerenciaMaterna: End Property
Public Property Get HerenciaFinal() As tResultado: HerenciaFinal = mHerenciaFinal: End Property
Public Property Get TipoHerencia() As Byte: TipoHerencia = mTipoHerencia: End Property
Public Property Get NotaHerencia() As String: NotaHerencia = mNotaHerencia: End Property

Public Property Get CicloActual() As Byte: CicloActual = mCicloActual: End Property
Public Property Get PinaculoActual() As Byte: PinaculoActual = mPinaculoActual: End Property
Public Property Get DesafioActual() As Byte: DesafioActual = mDesafioActual: End Property



' ============================================================
'   MÉTODOS PÚBLICOS
' ============================================================

' Cargar datos base
Public Sub CargarDatos(Nombre As String, Ape1 As String, Ape2 As String, FechaNac As Date)
    mNombre = Nombre
    mApe1 = Ape1
    mApe2 = Ape2
    mFechaNac = FechaNac
End Sub

' Método maestro
Public Sub CalcularTodo()
    CalcularFon
    CalcularValores
    CalcularNumerosPrincipales
    CalcularNumerosAdicionales
    CalcularHerencias
    CalcularCiclos
    CalcularPinaculos
    CalcularDesafios
End Sub



' ============================================================
'   MÉTODOS PRIVADOS (vacíos por ahora)
' ============================================================

Private Sub CalcularFon()
End Sub

Private Sub CalcularValores()
End Sub

Private Sub CalcularNumerosPrincipales()
End Sub

Private Sub CalcularNumerosAdicionales()
End Sub

Private Sub CalcularHerencias()
End Sub

Private Sub CalcularCiclos()
End Sub

Private Sub CalcularPinaculos()
End Sub

Private Sub CalcularDesafios()
End Sub

