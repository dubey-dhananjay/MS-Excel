Function IFS(ParamArray args() As Variant) As Variant
    
    Dim i As Long
    
    For i = LBound(args) To UBound(args) Step 2
        If args(i) Then
            IFS = args(i + 1)
            Exit Function
        End If
        
    Next i
    
    IFS = ""
    
End Function
