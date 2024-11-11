Public Function IsAllTrue(blnArray As Variant) As Boolean

    'PURPOSE: Check if a boolean Array has all its values as True
    
    Dim blnValue As Variant
    
    IsAllTrue = True
    
    For Each blnValue In blnArray
    
        If VarType(blnValue) = vbBoolean Then
            If blnValue <> True Then
                IsAllTrue = False
                Exit Function
            End If
        Else
            IsAllTrue = False
            Exit Function
        End If
        
    Next blnValue

End Function
