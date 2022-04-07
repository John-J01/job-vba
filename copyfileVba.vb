Sub CopyFileOnVBA()
Dim filePath As String
Row = WorksheetFunction.CountA(ThisWorkbook.Worksheets(1).Range("A:A"))
For CountRow = 1 To Row
'"C:\Users\admin8420\Pictures\Saved Pictures\Screenshot_1.png"
FileCopy "C:\Users\admin8420\Pictures\Saved Pictures\" & ThisWorkbook.Worksheets(1).Range("A" & CountRow).Value, _
    "C:\Users\admin8420\Pictures\Saved Pictures\new\" & ThisWorkbook.Worksheets(1).Range("A" & CountRow).Value
Next CountRow
End Sub
