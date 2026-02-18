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
    Dim cu As Integer
    Dim Gen As String
    Dim i As Integer
    
    
    arrFiles = Array("G_EquivalenciasCA", "G_EquivalenciasCA-IB", "G_EquivalenciasCA-VA", _
                     "G_EquivalenciasES", "G_EquivalenciasGL", "G_EquivalenciasEU", _
                     "G_EquivalenciasPT-EU", "G_EquivalenciasPT-BR", "G_EquivalenciasEN-GB", _
                     "G_EquivalenciasFR", "G_EquivalenciasEN-US", "G_EquivalenciasEN-US-AF")



    For Each tbl In arrFiles
        Set rsOri = CurrentDb.OpenRecordset("select * from [" & CStr(tbl) & "]")
        Set rsOut = CurrentDb.OpenRecordset("SELECT * FROM tbmEquivNombre_3")
        cu = rsOri.Fields.Count - 1
        
        Debug.Print tbl
        
        While Not rsOri.EOF
            DoEvents
            
            idOrigen = rsOri.Fields(0).Name
            nomOri = rsOri.Fields(0).Value
                        
            Gen = rsOri.Fields(cu).Value
            For i = 1 To cu - 1
            
                idEquiv = rsOri.Fields(i).Name
                NomEquiv = rsOri.Fields(i).Value
                
                rsOut.AddNew
                
                rsOut!NombreOriginal = nomOri
                rsOut!IdiomaOriginal = LCase(idOrigen)
                rsOut!NombreEquivalente = NomEquiv
                rsOut!IdiomaEquivalente = LCase(idEquiv)
                
                rsOut!genero = Gen
                
                If idEquiv <> "en-us-af" Then
                    rsOut!Activo = True
                End If
                
                rsOut.Update
            Next i
            rsOri.MoveNext
        Wend
    Next


End Sub
