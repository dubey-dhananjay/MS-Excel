Function TEXTJOIN(delimiter As String, ignore_empty As Boolean, rng As Range) As String
    
    Dim cell As Range
    
    For Each cell In rng
        If Not ignore_empty Or cell.Value <> "" Then
            TEXTJOIN = TEXTJOIN & cell.Value & delimiter
        End If
    Next cell
    
    TEXTJOIN = Left(TEXTJOIN, Len(TEXTJOIN) - Len(delimiter))

End Function
