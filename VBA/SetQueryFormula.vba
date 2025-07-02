Sub SetQueryFormula(queryName As String, value As Variant)

    Dim qry     As WorkbookQuery: Set qry = ThisWorkbook.Queries(queryName)
    Dim formula As String
    Dim i       As Long
    
    Select Case VarType(value)
    Case vbString
    
        formula = Replace(value, """", """""") 'Substitui todas as aspas duplas por 2 aspas duplas, para não ferrar o texto no Power Query
        formula = """" & formula & """"
    
    Case vbDate
    
        formula = "#date(" & year(value) & "," & month(value) & "," & day(value) & ")"
    
    Case vbArray + vbByte
    
        formula = "{" & value(0)
        For i = 1 To UBound(value)
            formula = formula & "," & value(i)
        Next i
        formula = formula & "}"
        
    End Select
    
    qry.formula = formula
    
End Sub
