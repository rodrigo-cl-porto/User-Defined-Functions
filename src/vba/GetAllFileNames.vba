Public Function GetAllFileNames(FolderPath As String, Optional fileExt As String) As String()

    'Get all file names of a given extension in a folder and in its subfolders

    Dim fso          As Object
    Dim oFolder      As Object
    Dim oFile        As Object
    Dim subFolder    As Object
    Dim i            As Long
    Dim NewUBound    As Long
    Dim colFolders   As New Collection
    Dim colFiles     As New Collection
    Dim arrFiles()   As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFolder = fso.GetFolder(FolderPath)
    
    colFolders.Add oFolder          'start with this folder
    
    Do While colFolders.Count > 0      'process all folders
        
        Set oFolder = colFolders(1)    'get a folder to process
        colFolders.Remove 1            'remove item at index 1
    
        For Each oFile In oFolder.Files
            If fileExt = vbNullString Then
                colFiles.Add oFile.name
            ElseIf InStr(1, fso.GetExtensionName(oFile.name), fileExt, vbTextCompare) Then
                colFiles.Add oFile.name
            End If
        Next oFile

        'add any subfolders to the collection for processing
        For Each subFolder In oFolder.subFolders
            colFolders.Add subFolder
        Next subFolder
        
    Loop
    
    'Converting collection into an array
    ReDim arrFiles(0 To colFiles.Count - 1)
    For i = 1 To colFiles.Count
        arrFiles(i - 1) = colFiles(i)
    Next i
    
    'Return array of file names
    GetAllFileNames = arrFiles

End Function
