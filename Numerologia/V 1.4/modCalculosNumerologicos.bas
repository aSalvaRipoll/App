Attribute VB_Name = "modCalculosNumerologicos"

Option Compare Database
Option Explicit

Private mSumaVocales As Integer
Private mSumaConsonantes As Integer
Private mSumaLetras As Integer

Public Sub CalcularResultado(ByRef R As clsResultado, ByRef P As clsPersona, ByRef f As clsFonetica)

    ' CÁLCULOS BASE DE LETRAS
    CargarAcumuladoresNombre f.NombreConvertido, f.sistema

    ' PRINCIPALES
    R.NumeroCaminoVida = CalcularCaminoVida(P)
    R.NumeroDestino = ReducirSimbolico(mSumaLetras)
    R.NumeroAlma = ReducirSimbolico(mSumaVocales)
    R.NumeroPersonalidad = ReducirSimbolico(mSumaConsonantes)

    ' DERIVADOS
    R.NumeroMadurez = CalcularMadurez(R)

End Sub


Public Sub CargarAcumuladoresNombre(nombre As String, sistema As String)

    Dim i As Integer
    Dim letra As String
    Dim valor As Byte

    mSumaVocales = 0
    mSumaConsonantes = 0
    mSumaLetras = 0
    
    For i = 1 To Len(nombre)
        letra = Mid(nombre, i, 1)

        ' Solo letras A-Z
        If letra Like "[A-Za-z]" Then

            valor = ConvertirLetraANumero(letra, sistema)

            If EsVocal(letra) Then
                mSumaVocales = mSumaVocales + valor
            Else
                mSumaConsonantes = mSumaConsonantes + valor
            End If
            mSumaLetras = mSumaLetras + valor
        End If
    Next i

End Sub

Private Function EsVocal(letra As String) As Boolean
    
    EsVocal = letra Like "[AEIOUÜaeiouü]"

End Function


Public Function CalcularCaminoVida(P As clsPersona) As String
    Dim anioRed As Long
    Dim suma As Long

    ' Reducir el año primero
    anioRed = SumarDigitos(Year(P.FechaNacimiento))

    ' Sumar día + mes + año reducido
    suma = Day(P.FechaNacimiento) + Month(P.FechaNacimiento) + anioRed

    ' Aplicar reducción simbólica elegante
    CalcularCaminoVida = ReducirSimbolico(suma)
End Function


Public Function CalcularDestino(f As clsFonetica) As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularDestino = ReducirSimbolico(mSumaLetras)
End Function

Public Function CalcularAlma(f As clsFonetica) As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularAlma = ReducirSimbolico(mSumaVocales)
End Function

Public Function CalcularPersonalidad(f As clsFonetica) As String
    ' Se asume que CargarAcumuladoresNombre ya se ha llamado antes
    CalcularPersonalidad = ReducirSimbolico(mSumaConsonantes)
End Function


Public Function CalcularMadurez(R As clsResultado) As String
    Dim cv As String, dest As String
    Dim rcv As Long, rdest As Long

    cv = R.NumeroCaminoVida
    dest = R.NumeroDestino

    rcv = CLng(Split(cv, "/")(UBound(Split(cv, "/"))))
    rdest = CLng(Split(dest, "/")(UBound(Split(dest, "/"))))

    CalcularMadurez = ReducirSimbolico(rcv + rdest)
End Function

'=====================================================================================










'Public Function SumarVocales(nombre As String, sistema As String) As Long
'    Dim i As Integer, letra As String
'    For i = 1 To Len(nombre)
'        letra = Mid(nombre, i, 1)
'        If letra Like "[AEIOUÜaeiouü]" Then
'            SumarVocales = SumarVocales + ConvertirLetraANumero(letra, sistema)
'        End If
'    Next i
'End Function
'
'Public Function SumarConsonantes(nombre As String, sistema As String) As Long
'    Dim i As Integer, letra As String
'    For i = 1 To Len(nombre)
'        letra = Mid(nombre, i, 1)
'        If letra Like "[A-Z]" And Not letra Like "[AEIOUÜ]" Then
'            SumarConsonantes = SumarConsonantes + ConvertirLetraANumero(letra, sistema)
'        End If
'    Next i
'End Function
'
'Public Function SumarTotal(nombre As String, sistema As String) As Long
'    SumarTotal = SumarVocales(nombre, sistema) + SumarConsonantes(nombre, sistema)
'End Function
'
