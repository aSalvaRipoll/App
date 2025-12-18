Attribute VB_Name = "modConstantesNumerologia"
Option Compare Database
Option Explicit

' ============================================================================
' Proyecto:     Sistema de Numerología Tradicional y Fonético
' Módulo:       modConstantesNumerologia
' Descripción:  Constantes globales para el sistema de numerología
' Autor:        Alba Salvá
' Fecha:        Diciembre 2025
' ============================================================================

' Números básicos
Public Const NUM_MIN As Integer = 1
Public Const NUM_MAX As Integer = 9

' Números Maestros
Public Const NUM_MAESTRO_11 As Integer = 11
Public Const NUM_MAESTRO_22 As Integer = 22
Public Const NUM_MAESTRO_33 As Integer = 33
Public Const NUM_MAESTRO_44 As Integer = 44

' Números Kármicos
Public Const NUM_KARMICO_13 As Integer = 13
Public Const NUM_KARMICO_14 As Integer = 14
Public Const NUM_KARMICO_16 As Integer = 16
Public Const NUM_KARMICO_19 As Integer = 19

' Valores numéricos de letras (Pitagórico)
Public Const VALOR_A As Integer = 1
Public Const VALOR_B As Integer = 2
Public Const VALOR_C As Integer = 3
Public Const VALOR_D As Integer = 4
Public Const VALOR_E As Integer = 5
Public Const VALOR_F As Integer = 6
Public Const VALOR_G As Integer = 7
Public Const VALOR_H As Integer = 8
Public Const VALOR_I As Integer = 9
Public Const VALOR_J As Integer = 1
Public Const VALOR_K As Integer = 2
Public Const VALOR_L As Integer = 3
Public Const VALOR_M As Integer = 4
Public Const VALOR_N As Integer = 5
Public Const VALOR_O As Integer = 6
Public Const VALOR_P As Integer = 7
Public Const VALOR_Q As Integer = 8
Public Const VALOR_R As Integer = 9
Public Const VALOR_S As Integer = 1
Public Const VALOR_T As Integer = 2
Public Const VALOR_U As Integer = 3
Public Const VALOR_V As Integer = 4
Public Const VALOR_W As Integer = 5
Public Const VALOR_X As Integer = 6
Public Const VALOR_Y As Integer = 7
Public Const VALOR_Z As Integer = 8

' Caracteres especiales españoles
Public Const VALOR_ENE As Integer = 5  ' Ñ
Public Const VALOR_CEDILLA As Integer = 3  ' Ç

' Vocales
Public Const VOCALES As String = "AEIOUÁÉÍÓÚÄËÏÖÜ"

' Consonantes (incluyendo Y cuando no es vocal)
Public Const CONSONANTES As String = "BCDFGHJKLMNPQRSTVWXYZÑÇ"

' Letras mudas
Public Const LETRA_H_MUDA As String = "H"
Public Const LETRA_U_MUDA_CONTEXTO As String = "QU,GU"  ' Contextos donde U es muda

' Rutas de archivos
Public Const RUTA_INTERPRETACIONES As String = "Interpretaciones\"
Public Const EXTENSION_MARKDOWN As String = ".md"

' Tipos de cálculo
Public Enum tipoCalculo
    CaminoVida = 1
    Destino = 2
    Alma = 3
    Personalidad = 4
    Madurez = 5
    AnoPersonal = 6
    MesPersonal = 7
    DiaPersonal = 8
    Ciclo1 = 9
    Ciclo2 = 10
    Ciclo3 = 11
    Ciclo4 = 12
    Pinaculo1 = 13
    Pinaculo2 = 14
    Pinaculo3 = 15
    Pinaculo4 = 16
    Desafio1 = 17
    Desafio2 = 18
    Desafio3 = 19
    Desafio4 = 20
    NumeroExpresion = 21
    NumeroPoder = 22
    numeroFaltante = 23
    NumeroDominante = 24
    PrimeraLetra = 25
    PrimeraVocal = 26
    PrimeraConsonante = 27
    RespuestaSubconsciente = 28
    PlanoExpresion = 29
End Enum

' Mensajes de error
Public Const ERR_FECHA_INVALIDA As String = "Fecha de nacimiento inválida"
Public Const ERR_NOMBRE_VACIO As String = "El nombre no puede estar vacío"
Public Const ERR_NUMERO_INVALIDO As String = "Número fuera de rango válido"
Public Const ERR_ARCHIVO_NO_ENCONTRADO As String = "Archivo de interpretación no encontrado"
Public Const ERR_CALCULO_FALLIDO As String = "Error en el cálculo numerológico"

' Mensajes informativos
Public Const MSG_NUMERO_MAESTRO As String = "Número Maestro detectado"
Public Const MSG_NUMERO_KARMICO As String = "Número Kármico detectado"
Public Const MSG_CALCULO_EXITOSO As String = "Cálculo completado exitosamente"
