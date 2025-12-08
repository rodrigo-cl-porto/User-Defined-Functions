Function RangeToHtml(rng As Range) As String

    Dim fso      As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts       As Object
    Dim TempFile As String
    Dim TempWB   As Workbook

    On Error GoTo Err
    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
    
        With .Cells(1)
            .PasteSpecial Paste:=8
            .PasteSpecial xlPasteValues, , False, False
            .PasteSpecial xlPasteFormats, , False, False
            .Select
            Application.CutCopyMode = False
        End With
        
        On Error Resume Next
        With .DrawingObjects
            .Visible = True
            .Delete
        End With
        On Error GoTo Err
        
    End With

    ActiveSheet.Range("A:D").Columns.AutoFit

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
        SourceType:=xlSourceRange, _
            Filename:=TempFile, _
            Sheet:=TempWB.Sheets(1).Name, _
            Source:=TempWB.Sheets(1).UsedRange.Address, _
            HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangeToHtml = ts.readall
    ts.Close
    RangeToHtml = Replace(RangeToHtml, "align=center x:publishsource=", "align=left x:publishsource=")

Err:

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the html file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing

End Function
