Sub FileListingAllFolder()
    
' Open folder selection
' Open folder selection

With Application.FileDialog(msoFileDialogFolderPicker)
    .Title = "Select a Folder"
    .AllowMultiSelect = False
    If .Show <> -1 Then GoTo NextCode
    pPath = .SelectedItems(1)
        If Right(pPath, 1) <> "\" Then
            pPath = pPath & "\"
        End If
End With

Application.WindowState = xlMinimized
Application.ScreenUpdating = False

    Workbooks.Add ' create a new workbook for the file list
    ' add headers
    ActiveSheet.Name = "ListOfFiles"
    With Range("A2")
        .Formula = "Folder contents:"
        .Font.Bold = True
        .Font.Size = 12
    End With
    Range("A3").Formula = "File Name:"
    Range("B3").Formula = "File Size:"
    Range("C3").Formula = "File Type:"
    Range("D3").Formula = "Date Created:"
    Range("E3").Formula = "Date Last Accessed:"
    Range("F3").Formula = "Date Last Modified:"
    Range("A3:F3").Font.Bold = True

    Worksheets("ListOfFiles").Range("A1").Value = pPath
    
        Range("A1").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .Color = -16776961
            .TintAndShade = 0
        End With
        Selection.Font.Bold = True
    
    ListFilesInFolder Worksheets("ListOfFiles").Range("A1").Value, True
    ' list all files included subfolders

    Range("A3").Select
    
    Lastrow = Range("A1048576").End(xlUp).Row
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select

    ActiveWorkbook.Worksheets("ListOfFiles").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ListOfFiles").Sort.SortFields.Add Key:=Range( _
        "B4:B" & Lastrow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ListOfFiles").Sort
        .SetRange Range("A3:F" & Lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A1").Select
    
Cells.Select
Cells.EntireColumn.AutoFit
Columns("A:A").Select
Selection.ColumnWidth = 100
Range("A1").Select
   
NextCode:
MsgBox "No files Selected!!"

End Sub

Sub ListFilesInFolder(SourceFolderName As String, IncludeSubfolders As Boolean)
' lists information about the files in SourceFolder
Dim FSO As Scripting.FileSystemObject
Dim SourceFolder As Scripting.Folder, SubFolder As Scripting.Folder
Dim FileItem As Scripting.File
Dim r As Long
    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    r = Range("A1000").End(xlUp).Row + 1
    For Each FileItem In SourceFolder.Files
        ' display file properties
        Cells(r, 1).Formula = FileItem.Path & FileItem.Name
        Cells(r, 2).Formula = (FileItem.Size / 1048576)
            Cells(r, 2).Value = Format(Cells(r, 2).Value, "##.##") & " MB"
        Cells(r, 3).Formula = FileItem.Type
        Cells(r, 4).Formula = FileItem.DateCreated
        Cells(r, 5).Formula = FileItem.DateLastAccessed
        Cells(r, 6).Formula = FileItem.DateLastModified
        ' use file methods (not proper in this example)

        r = r + 1 ' next row number
    Next FileItem
    If IncludeSubfolders Then
        For Each SubFolder In SourceFolder.SubFolders
            ListFilesInFolder SubFolder.Path, True
        Next SubFolder
    End If
    Columns("A:F").AutoFit
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
    ActiveWorkbook.Saved = True
End Sub

