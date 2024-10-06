Function CONCAT(ParamArray args() As Variant) As String
    
    Dim i As Long
    Dim temp As String
    
    For i = LBound(args) To UBound(args)
        If IsArray(args(i)) Then
            If TypeName(args(i)) = "Range" Then
                temp = ""
                For Each cell In args(i)
                    temp = temp & cell.Value
                Next cell
            Else
                temp = Join(args(i), "")
            End If
            CONCAT = CONCAT & temp
        Else
            CONCAT = CONCAT & args(i)
        End If
    
    Next i
    
End Function
