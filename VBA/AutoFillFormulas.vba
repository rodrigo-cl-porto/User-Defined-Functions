Public Sub AutoFillFormulas(rng As Range, Optional UseLastCellAsRef As Boolean = False)

    Dim RefCell As Range
    
    If rng Is Nothing Then Exit Sub 'If range is nothing, don't do anything
    If rng.Count = 1 Then Exit Sub  'if range has only 1 cell, don't do anything
    
    If RangeHasAnyFormula(rng) Then
    
        With rng.SpecialCells(xlCellTypeFormulas)
            If Not UseLastCellAsRef Then
                Set RefCell = .Cells(1)
            Else
                Set RefCell = .Cells(.Count)
            End If
        End With
        
        rng.FormulaR1C1 = RefCell.FormulaR1C1
        
    End If

End Sub
