'Last update 24/12/2021

''----------------PRINT PAPER------------------

Sub printf()
ActiveWindow.SelectedSheets.PrintOut
Call loop_name
End Sub


''--------------- CREATE NEW SHEET OR UPDATE NEW ROWS-----------------

Sub loop_name()
Set TEM = ThisWorkbook.Sheets("TEM")
Set khsx = ThisWorkbook.Sheets("KHSX")
Set MASP = TEM.Range("T16")
stt = TEM.Range("b40").Value

 Dim name, sheet As Worksheet
test = 0
For Each name In Worksheets
    If name.name = MASP Then
        test = test + 1
        Set sheet = name
    End If
Next name
'Check double
If test = 0 Then
'if not exist
    Sheets("Template").Select
    'ActiveWindow.SmallScroll Down:=-20
    Cells.Select
    Selection.Copy
        Sheets.Add
        ActiveSheet.name = MASP
        ActiveSheet.Range("A1").Select
         ActiveSheet.Paste
        ActiveSheet.Range("A2").Value = TEM.Range("J28").Value
        ActiveSheet.Range("B2").Value = TEM.Range("D36").Value & "*" & TEM.Range("I36").Value
        ActiveSheet.Range("C2").Value = Application.WorksheetFunction.VLookup(stt, khsx.Range("A:bQ"), 41, 0)
        ActiveSheet.Range("B4").Value = TEM.Range("D41").Value
        If Not IsEmpty(TEM.Range("E33")) Then
        ActiveSheet.Range("E4").Value = TEM.Range("E35").Value
        End If
        ActiveSheet.Select
        
   Else
   'Neu da co
        N = Application.WorksheetFunction.CountA(sheet.Range("B:B"))
                sheet.Range("B" & N + 1).Value = TEM.Range("D41").Value
                 If Not IsEmpty(TEM.Range("E35")) Then
                       sheet.Range("E" & N + 1).Value = TEM.Range("E35").Value
                 End If
                 sheet.Select
    End If
End Sub

'----------------CHECK SUM---------------

Sub tong_auto()
  Set tong = ThisWorkbook.Sheets("TONG")
  tong.Select
  tong.Range("B3:K80").ClearContents
For Each name In Worksheets
 If name.name = "KHSX" Or name.name = "TEM" Or name.name = "TONG" Or name.name = "Template" Then
 Else
        N = Application.WorksheetFunction.CountA(tong.Range("B:B"))
        tong.Range("B" & N + 1).Value = name.name
        tong.Range("C" & N + 1).Value = name.Range("C2").Value
        tong.Range("D" & N + 1).Value = name.Range("F2").Value
        tong.Range("F" & N + 1).Value = name.Range("E2").Value
        tong.Range("G" & N + 1).Value = name.Range("H2").Value
        tong.Range("H" & N + 1).Value = name.Range("J2").Value
        tong.Range("E" & N + 1).Value = name.Range("E2").Value
  End If
  'tong.Range("B4:K80").Selection.NumberFormat = "0"
Next name
End Sub



'' ----------------CREATE DELEVERY RECORDS--------------


Sub Phieu_BG()
Set bg = ThisWorkbook.Sheets("PHIEU BG")
Dim Actsheet As Worksheet
num = 0
For Each name In ThisWorkbook.Worksheets
    If name.name = bg.Range("H4") & "-" & bg.Range("G7") Then
        Set Actsheet = name
        num = num + 1
    End If
Next name
If num = 0 Then
MsgBox "Don hang chua chay", vbExclamation
End
End If

If bg.Range("H4") & "-" & bg.Range("G7") = bg.Range("H3").Value Then
bg.Range("G12").Value = bg.Range("G12").Value + 1
Else
bg.Range("G12").Value = 1
End If
so_phieu = bg.Range("G12").Value
'================
'check stt stt = 1 tuong ung 20
'================
bg.Range("C15:D35").ClearContents
bg.Range("f15:G35").ClearContents
bg.Rows("15:35").EntireRow.Hidden = False
Actsheet.Select
'neu so phieu = 1
If so_phieu = 1 Then
Actsheet.Range("c4:A23").Copy
sumValue = Application.WorksheetFunction.Sum(Actsheet.Range("C4:C23")) 'check con thieu
Else
Actsheet.Range("c" & (4 + (so_phieu - 1) * 20) & ":" & "A" & (23 + (so_phieu - 1) * 20)).Copy
sumValue = Application.WorksheetFunction.Sum(Actsheet.Range("C4:C" & (23 + (so_phieu - 1) * 20))) 'check con thieu
End If
bg.Range("B108").PasteSpecial Paste:=xlPasteValues
For i = 108 To 128
If bg.Range("D" & i).Value = 0 Then
bg.Rows(i & ":" & i).EntireRow.Delete
End If
Next i

bg.Range("D108:C128").Copy
bg.Range("c15").PasteSpecial Paste:=xlPasteValues

bg.Range("B108:B128").Copy
bg.Range("c15").PasteSpecial Paste:=xlPasteValues

For i = 15 To 34
If bg.Range("D" & i).Value = 0 Then bg.Rows(i & ":" & i).EntireRow.Hidden = True

Next i
'check trung lap
bg.Range("H3").Value = bg.Range("H4") & "-" & bg.Range("G7")
bg.Range("B94:D170").ClearContents
bg.Select
ActiveWindow.SelectedSheets.PrintPreview
Call sb_Copy_Save_Worksheet_As_Workbook
End Sub

Sub sb_Copy_Save_Worksheet_As_Workbook()
nameWB = ThisWorkbook.name
Set strFile = ThisWorkbook.Sheets("PHIEU BG").Range("H5")
Set strFolder = ThisWorkbook.Sheets("PHIEU BG").Range("H6")
If Dir(strFolder) <> "" Then
     Dim wb As Workbook
    Set wb = Workbooks.Add
    Workbooks(nameWB).Sheets("PHIEU BG").Copy Before:=wb.Sheets(1)
   File = (strFolder & strFile)
   Application.DisplayAlerts = False
    wb.SaveAs strFolder & strFile
    wb.Close SaveChanges:=True
   Else
   MsgBox "khong co folder nay"
   End If
    'wb.Close SaveChanges:= True Filename:= file
End Sub

Sub only_print()
ActiveWindow.SelectedSheets.PrintPreview
End Sub
