'Last update 18/03/2022 

''----------------PRINT PAPER------------------

Sub printf()
ActiveWindow.SelectedSheets.PrintPreview
Call insert_new_row
End Sub

''--------------- CREATE NEW ROWS-----------------


Sub insert_new_row()
Set Sh_Tem = ThisWorkbook.Sheets("TEM IN")
Set Sh_Tong_Hop = ThisWorkbook.Sheets("TONG_HOP")
        CountRows = Application.WorksheetFunction.CountA(Sh_Tong_Hop.Range("B:B"))
                    Sh_Tong_Hop.Range("a" & CountRows + 1).Value = Sh_Tem.Range("B39").Value
                    Sh_Tong_Hop.Range("B" & CountRows + 1).Value = Sh_Tem.Range("j28").Value
                    Sh_Tong_Hop.Range("c" & CountRows + 1).Value = Sh_Tem.Range("D31").Value
                    Sh_Tong_Hop.Range("d" & CountRows + 1).Value = Sh_Tem.Range("k33").Value
                    Sh_Tong_Hop.Range("e" & CountRows + 1).Value = Sh_Tem.Range("b37").Value
                    Sh_Tong_Hop.Range("j" & CountRows + 1).Value = Sh_Tem.Range("D39").Value
                    Sh_Tong_Hop.Range("g" & CountRows + 1).Value = Sh_Tem.Range("H39").Value
                    Sh_Tong_Hop.Range("h" & CountRows + 1).Value = Sh_Tem.Range("f37").Value
                    Sh_Tong_Hop.Range("I" & CountRows + 1).Value = Sh_Tem.Range("L37").Value
                    Sh_Tong_Hop.Range("f" & CountRows + 1).Value = Sh_Tem.Range("B41").Value
Call DeleteShapes
End Sub
Sub newData()
Set tonghop = ThisWorkbook.Sheets("TONG_HOP")
Set data = ThisWorkbook.Sheets("data")
msp = ThisWorkbook.Sheets("PHIEU BG").Range("D12").Value
' clear contents
data.Range("A1:V1000").ClearContents
tonghop.Range("A1:V1").AutoFilter _
 Field:=1, _
 Criteria1:=msp
  ' copy to new sheet
 tonghop.Range("A1:V1000").SpecialCells(xlCellTypeVisible).Copy
    data.Range("A1").PasteSpecial
'no filter
 tonghop.ShowAllData
 ThisWorkbook.Sheets("PHIEU BG").Select
End Sub
Sub Phieu_BG()
Call newData
Set bg = ThisWorkbook.Sheets("PHIEU BG")
Set Actsheet = ThisWorkbook.Sheets("Data")
If bg.Range("H4") & "-" & bg.Range("G12") = bg.Range("H3").Value Then
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
Actsheet.Range("f2:g21").Copy
sumValue = Application.WorksheetFunction.Sum(Actsheet.Range("C4:C23")) 'check con thieu
Else
Actsheet.Range("f" & (2 + (so_phieu - 1) * 20) & ":" & "g" & (21 + (so_phieu - 1) * 20)).Copy
sumValue = Application.WorksheetFunction.Sum(Actsheet.Range("C4:C" & (21 + (so_phieu - 1) * 20))) 'check con thieu
End If
bg.Range("B108").PasteSpecial Paste:=xlPasteValues
bg.Range("D108:C128").Copy
bg.Range("c15").PasteSpecial Paste:=xlPasteValues

bg.Range("B108:C128").Copy
bg.Range("c15").PasteSpecial Paste:=xlPasteValues

For i = 15 To 34
If bg.Range("D" & i).Value = 0 Then bg.Rows(i & ":" & i).EntireRow.Hidden = True

Next i
'check trung lap
 bg.Range("H3").Value = bg.Range("H4") & "-" & bg.Range("G12")
 bg.Range("B94:D170").ClearContents
 ThisWorkbook.Sheets("PHIEU BG").Select
ActiveWindow.SelectedSheets.PrintPreview
'Call sb_Copy_Save_Worksheet_As_Workbook
Call DeleteShapes
End Sub

'------------QR CODE------------------


Function GenerateQR(qrcode_value As String)
Set TEM = ThisWorkbook.Sheets("TEM IN")
    Dim URL As String
    Dim My_Cell As Range
    Set My_Cell = Application.Caller
    URL = "https://chart.googleapis.com/chart?chs=90x90&&cht=qr&chl=" & qrcode_value
    Debug.Print URL
    On Error Resume Next
      TEM.Pictures("My_QR_CODE_" & My_Cell.Address(False, False)).Delete
    On Error GoTo 0
    TEM.Pictures.Insert(URL).Select
    With Selection.ShapeRange(1)
     .name = "My_QR_CODE_" & My_Cell.Address(False, False)
     .Left = My_Cell.Left + 10
     .Top = My_Cell.Top + 15
    End With
    GenerateQR = " "
End Function

'-----------------DELETE SHAPES-----------

Sub DeleteShapes()
Set TEM = ThisWorkbook.Sheets("TEM IN")
Dim shape As Excel.shape
For Each shape In TEM.Shapes
If shape.name = "LOGO" Or shape.name = "My_QR_CODE_D19" Or shape.name = "Button 2" Or shape.name = "Button 1" Or shape.name = "Button 3" Then
' do no thing
Else
shape.Delete
End If
Next
End Sub


'----------------GET THE DATA OF THE FILE---------------

Sub laykhsx()
Set khsx = ThisWorkbook.Sheets("KHSX")
  Dim wb As Workbook
    Dim strFolder As String
    Dim strFile As String
    strFile = ThisWorkbook.Sheets("KHSX").Range("B2").Value & ".xlsx"
        strFolder = ThisWorkbook.Sheets("KHSX").Range("B1").Value
        
    strFileExists = Dir(strFolder & strFile)
   If strFileExists = "" Then
        MsgBox "Không có file kê hoach nay"
    Else
    khsx.Range("A8", khsx.Range("AA200").End(xlDown)).ClearContents
            Set wb = Workbooks.Open(strFolder & strFile, UpdateLinks:=False, ReadOnly:=True)
            'CHECK DONE
wb.Sheets("LENH SAN XUAT").Range("A7", wb.Sheets("LENH SAN XUAT").Range("AA7").End(xlDown)).Copy
khsx.Range("A7").PasteSpecial Paste:=xlPasteValues
Application.DisplayAlerts = False
wb.Close SaveChanges:=True
Call DeleteShapes
End If
End Sub
