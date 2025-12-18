Option Compare Database
Option Explicit

' =============================================================================
' Módulo: modConstantesNumerologia
' Descripción: Constantes y enumeraciones para el sistema numerológico
' Autor: Sistema de Numerología
' Fecha: 2025
' Versión: 2.0 - Añadido soporte para Día de Nacimiento
' =============================================================================

' Enumeración para tipos de interpretación
Public Enum TipoInterpretacion
    tiCaminoVida = 1
    tiDestino = 2
    tiAlma = 3
    tiPersonalidad = 4
    tiMadurez = 5
    tiSinastria = 6
    tiDiaNacimiento = 7      ' NUEVO: Día de Nacimiento
End Enum

' Enumeración para tipos de letras
Public Enum TipoLetra
    tlVocal = 1
    tlConsonante = 2
    tlEspecial = 3
    tlYContextual = 4
End Enum

' Enumeración para números maestros
Public Enum NumeroMaestro
    nmNinguno = 0
    nmOnce = 11
    nmVeintidos = 22
    nmTreintaTres = 33
    nmCuarentaCuatro = 44
End Enum

' Constantes para validación
Public Const VOCALES_STANDARD As String = "AEIOU"
Public Const CONSONANTES_STANDARD As String = "BCDFGHJKLMNPQRSTVWXZ"
Public Const LETRA_ESPECIAL_EÑE As String = "Ñ"
Public Const LETRA_ESPECIAL_CEDILLA As String = "Ç"
Public Const LETRA_CONTEXTUAL_Y As String = "Y"

' Constantes para valores numerológicos
Public Const VALOR_EÑE As Integer = 5  ' Ñ = 14 → 1+4 = 5
Public Const VALOR_CEDILLA As Integer = 8  ' Ç = 8
Public Const VALOR_Y_VOCAL As Integer = 9  ' Y como vocal = 9
Public Const VALOR_Y_CONSONANTE As Integer = 7  ' Y como consonante = 7

' Constantes para mensajes
Public Const MSG_ERROR_LETRA_INVALIDA As String = "La letra proporcionada no es válida"
Public Const MSG_ERROR_CONVERSION As String = "Error al convertir la letra a número"
Public Const MSG_LETRA_NO_ENCONTRADA As String = "Letra no encontrada en la tabla de conversión"
