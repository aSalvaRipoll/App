Public Function DConcat(SQL As String, Optional Sep As String = ", ") As String
    Dim rs As DAO.Recordset
    Dim tmp As String

    Set rs = CurrentDb.OpenRecordset(SQL)
    While Not rs.EOF
        tmp = tmp & rs.Fields(0).Value & Sep
        rs.MoveNext
    Wend
    rs.Close

    If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - Len(Sep))
    DConcat = tmp
End Function


Public Function DConcat( _
        ByVal Campo As String, _
        ByVal Tabla As String, _
        Optional ByVal Where As String = "", _
        Optional ByVal Sep As String = ", ") As String
    
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim tmp As String
    
    ' Construir SQL dinámico
    sql = "SELECT " & Campo & " FROM " & Tabla
    If Len(Where) > 0 Then
        sql = sql & " WHERE " & Where
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    While Not rs.EOF
        If Not IsNull(rs.Fields(0).Value) Then
            tmp = tmp & rs.Fields(0).Value & Sep
        End If
        rs.MoveNext
    Wend
    
    rs.Close
    
    ' Quitar separador final
    If Len(tmp) > 0 Then
        tmp = Left(tmp, Len(tmp) - Len(Sep))
    End If
    
    DConcat = tmp
End Function
