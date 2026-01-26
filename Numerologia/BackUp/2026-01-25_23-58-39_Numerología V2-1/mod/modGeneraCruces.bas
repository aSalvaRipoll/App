Attribute VB_Name = "modGeneraCruces"
Option Compare Database
Option Explicit

Sub GeneraCruces()

    Dim rsOri As DAO.Recordset
    Dim rsOut As DAO.Recordset
    Dim arrFiles As Variant
    Dim tbl As Variant
    Dim nomOri As String, idOrigen As String
    Dim NomEquiv As String, idEquiv As String
    
    Dim i As Integer
    
    
    arrFiles = Array("EquivalenciasCA", "EquivalenciasCA-IB", "EquivalenciasCA-VA", _
                     "EquivalenciasES", "EquivalenciasGL", "EquivalenciasEU", _
                     "EquivalenciasPT-EU", "EquivalenciasPT-BR", "EquivalenciasEN-GB", _
                     "EquivalenciasFR", "EquivalenciasEN-US")



    For Each tbl In arrFiles
        Set rsOri = CurrentDb.OpenRecordset("select * from [" & CStr(tbl) & "]")
        Set rsOut = CurrentDb.OpenRecordset("SELECT * FROM tbmEquivNombre_2")
        
        While Not rsOri.EOF
            DoEvents
            
            idOrigen = rsOri.Fields(0).Name
            nomOri = rsOri.Fields(0).Value
                        
            For i = 1 To 11
            
                idEquiv = rsOri.Fields(i).Name
                NomEquiv = rsOri.Fields(i).Value
                If idEquiv <> "en-us-af" Then
                    rsOut.AddNew
                    
                    rsOut!NombreOriginal = nomOri
                    rsOut!IdiomaOriginal = LCase(idOrigen)
                    rsOut!NombreEquivalente = NomEquiv
                    rsOut!IdiomaEquivalente = LCase(idEquiv)
                        
                    rsOut.Update
                End If
            Next i
            rsOri.MoveNext
        Wend
    Next


End Sub
