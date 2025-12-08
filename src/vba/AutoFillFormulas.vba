Private Function RangeHasAnyFormula(ByVal rng As Range) As Boolean

    Dim ErrNumber As Integer
    Dim ErrText   As String
    Dim Return    As Boolean

    If rng Is Nothing Then Exit Function

    On Error Resume Next
    rng.SpecialCells xlCellTypeFormulas
    ErrNumber = Err.Number
    ErrText = Err.Description
    On Error GoTo 0
    
    If ErrNumber = 0 Then
        Return = True
    ElseIf ErrText = "No cells were found." Then
        Return = False
    Else
        MsgBox "The following error occured: " & ErrNumber & vbLf & vbLf & ErrText, vbCritical + vbOKOnly, "Message Error"
    End If

    RangeHasAnyFormula = Return

End Function

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
