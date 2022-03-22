'Last update 05/05/2021 

'----------CREATE-----------------

Sub CreateData(NameSheet As String, ByVal key As MSForms.TextBox, rangeKey As String, keysheet As String)
Dim count As Integer
Dim nameSh As Worksheet
Set nameSh = ThisWorkbook.sheets(NameSheet)
'check same value
If Application.WorksheetFunction.CountIf(nameSh.Range(rangeKey & ":" & rangeKey), key.Value) > 0 And key.Value <> "" Then
        MsgBox "Du lieu da bi trung lap", vbInformation
ElseIf key.Value = "" Then
        MsgBox "Vui long nhap key "
Else
    Dim Row As Long
    'NHAP DU LIEU VAO SHEET
    Row = Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey))
    For count = 1 To Application.WorksheetFunction.CountA(nameSh.Rows(1))
    'check index

    If count = 1 Then
        If UserForm1.Controls(keysheet & 1).Value = "" Then 'check update
            If Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)) = 1 Then
                nameSh.Cells(Row + 1, count) = 1
            Else
                nameSh.Cells(Row + 1, count) = nameSh.Cells(Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)), count).Value + 1
            End If
            End If
        Else
            nameSh.Cells(Row + 1, count).Value = UserForm1.Controls(keysheet & count).Value
        End If
    Next count
    If keysheet = "P" Then
        Call clearDataSp
        Call UserForm1.ShowlistB
    End If
         MsgBox "Tao du lieu thanh cong"
End If
End Sub

Sub CreateDataKeyId(NameSheet As String, ByVal key As MSForms.TextBox, rangeKey As String, keysheet As String)

Dim count As Integer
Dim nameSh As Worksheet
Set nameSh = ThisWorkbook.sheets(NameSheet)
'check same value
If Application.WorksheetFunction.CountIf(nameSh.Range(rangeKey & ":" & rangeKey), key.Value) > 0 And key.Value <> "" Then
    MsgBox "Du lieu da bi trung lap", vbInformation
ElseIf key.Value = "" Then
    MsgBox "Vui long nhap key "
    Else
        Dim Row As Long
'NHAP DU LIEU VAO SHEET
        Row = Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey))
        For count = 1 To Application.WorksheetFunction.CountA(nameSh.Rows(1))
'check index
        If count = 1 Then
'If UserForm1.Controls(keysheet & 1).Value = "" Then 'check update
            If Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)) = 1 Then
                nameSh.Cells(Row + 1, count) = 1
            Else
                nameSh.Cells(Row + 1, count) = nameSh.Cells(Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)), count).Value + 1
'End If
            End If
        Else
            nameSh.Cells(Row + 1, count).Value = UserForm1.Controls(keysheet & count).Value
        End If
        Next count

        If keysheet = "P" Then
            Call clearDataSp
            Call UserForm1.ShowlistB
        End If
            MsgBox "Tao du lieu thanh cong"
End If
End Sub

'Application.InputBox
Sub CreateNoKey(NameSheet As String, ByVal key As MSForms.TextBox, rangeKey As String, keysheet As String)
Dim count As Integer
Dim nameSh As Worksheet
    Set nameSh = ThisWorkbook.sheets(NameSheet)
'check same value
If Application.WorksheetFunction.CountIf(nameSh.Range(rangeKey & ":" & rangeKey), key.Value) > 0 And key.Value <> "" Then
    MsgBox "Du lieu da bi trung lap", vbInformation
Else
    Dim Row As Long
'NHAP DU LIEU VAO SHEET
    Row = Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey))
    For count = 1 To Application.WorksheetFunction.CountA(nameSh.Rows(1))
'check index
    If count = 1 Then
        If UserForm1.Controls(keysheet & 1).Value = "" Then 'check update
            If Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)) = 1 Then
                nameSh.Cells(Row + 1, count) = 1
            Else
                nameSh.Cells(Row + 1, count) = nameSh.Cells(Application.WorksheetFunction.CountA(nameSh.Range(rangeKey & ":" & rangeKey)), count).Value + 1
            End If
        End If
    Else
        nameSh.Cells(Row + 1, count).Value = UserForm1.Controls(keysheet & count).Value
    End If

Next count
    MsgBox "DU LIEU DA DUOC TAO"
    If keysheet = "LG" Then
        Call clearDataLG
        Call UserForm1.ShowlistLG
    End If
    If keysheet = "LO" Then
        Call UserForm1.clearDataLO
        Call UserForm1.ShowlistLO
    End If
End If
End Sub


'--------------------DELETE--------------------------

Sub delete(nameSh As String, ByVal listb As MSForms.ListBox)
'DELETE ROW
If listb.ListIndex < 0 Then
    MsgBox "Vui long chon 1 ban ghi", vbInformation
    Exit Sub
End If
    Dim r As Integer
    r = listb.ListIndex + 2
    sheets(nameSh).Rows(r).delete
    MsgBox "du lieu da duoc xoa"
End Sub


'-----------------EXPORT----------------------------


Sub NewShPreview(nameSh As String)
'TAO SHEET MOI VA PRINTF
Dim news As Workbook
Set news = Workbooks.Add
    ThisWorkbook.sheets(nameSh).UsedRange.Copy news.sheets(1).Range("A1")
    UserForm1.Hide
    ActiveWindow.SelectedSheets.PrintPreview
End Sub

Sub NewShFrint(nameSh As String)
'TAO SHEET MOI VA PRINTF
Dim news As Workbook
Set news = Workbooks.Add
    ThisWorkbook.sheets(nameSh).UsedRange.Copy news.sheets(1).Range("A1")
    UserForm1.Hide
    ActiveWindow.SelectedSheets.PrintPreview
End Sub


'--------------SET VALUE-------------------------

Sub NewValue(name As String, rg As String, ByVal listb As MSForms.ComboBox)
''set value filter
    Dim wb As Worksheet
    Set wb = ThisWorkbook.sheets(name)
    Dim count As Integer
    listb.Clear 'set list box
For count = 2 To Application.WorksheetFunction.CountA(wb.Range(rg & ":" & rg))
        If count > 1 Then
            listb.AddItem ThisWorkbook.sheets(name).Range(rg & count).Value
        End If
Next count
End Sub

Sub NewValueFilter(name As String, rg As String, ByVal listb As MSForms.ComboBox, rgcheck As String, namesheetdest As String, rangedest As String)
''set value filter
    Dim wb As Worksheet
    Set wb = ThisWorkbook.sheets(name)
' set tim kiem gia tri
    Dim dest As Worksheet
    Set dest = ThisWorkbook.sheets(namesheetdest)
    Dim count, fil As Integer
    listb.Clear 'set list box
For count = 2 To Application.WorksheetFunction.CountA(wb.Range(rg & ":" & rg))
    If count > 1 And Application.WorksheetFunction.VLookup(wb.Range(rgcheck & count).Value, dest.Range(rangedest & ":" & "AA"), 2, 0) <> 0 Then
        listb.AddItem wb.Range(rg & count).Value
    End If
Next count
End Sub

Sub NewValueRange(name As String, ByVal listb As MSForms.ComboBox, rgcheck As String, listCk As MSForms.ComboBox, rgDest As String, rangeSource As String, rangeSearch As String)
''set value filter
    Dim wb As Worksheet
    Set wb = ThisWorkbook.sheets(name)
    Dim source As Worksheet
    Set source = ThisWorkbook.sheets("SOURCE")
    Dim count As Integer
        listb.Clear 'set list box
    For count = 2 To Application.WorksheetFunction.CountA(wb.Range(rgcheck & ":" & rgcheck))
If count > 1 And wb.Range(rgcheck & count).Value = listCk.Value Then
    If Application.WorksheetFunction.VLookup(wb.Range(rangeSearch & count).Value, source.Range(rangeSource & ":" & "AA"), 2, 0) <> 0 Then
        listb.AddItem wb.Range(rgDest & count).Value
    End If
End If
Next count
End Sub


'----------------- SHOW LIST--------------------

Sub ShowListBox(NameSheet As String, ByVal List As MSForms.ListBox)
    Dim nameSh As Worksheet
    Set nameSh = ThisWorkbook.sheets(NameSheet)
    last_row = Application.WorksheetFunction.CountA(nameSh.Range("A:A"))
If last_row = 1 Then last_row = 2
    With List
        .ColumnHeads = True
        .ColumnCount = Application.WorksheetFunction.CountA(nameSh.Rows(1))
        .RowSource = nameSh.name & "!A2:CA" & last_row
    End With
End Sub

'========== List Filter ==========

Sub showListFilter(NameSheet As String, NameNewSheet As String, ByVal F As MSForms.ComboBox, ByVal T As MSForms.TextBox, ByVal List As MSForms.ListBox)
'sheet moi de filter
    Dim NameNewSh As Worksheet
    Set NameNewSh = ThisWorkbook.sheets(NameNewSheet)
'sheet
    Dim nameSh As Worksheet
    Set nameSh = ThisWorkbook.sheets(NameSheet)
'filter du lieu
nameSh.AutoFilterMode = False
If F.Value = "ALL" Then
    nameSh.AutoFilterMode = False
ElseIf F.Value <> "All" And T.Value <> "" Then
    nameSh.UsedRange.AutoFilter Application.WorksheetFunction.Match(F.Value, nameSh.Range("A1:CA1"), 0), T.Value
End If
    NameNewSh.Cells.Clear
    nameSh.UsedRange.Copy
    NameNewSh.Range("A1").PasteSpecial xlPasteValues
    NameNewSh.Range("A1").PasteSpecial xlPasteFormats
    nameSh.AutoFilterMode = False
    last_row = Application.WorksheetFunction.CountA(nameSh.Range("A:A"))
    If last_row = 1 Then last_row = 2
With List
    .ColumnHeads = True
    .ColumnCount = Application.WorksheetFunction.CountA(nameSh.Rows(1))
    .RowSource = NameNewSh.name & "!A2:CA" & last_row
    .Font.Size = 8
End With

End Sub

'-----------------SHOW DATA---------------------------
Sub SortAz(nameSh As String, ByVal listb As MSForms.ComboBox)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.sheets(nameSh)
If listb.Value = "ALL" Or listb.Value = "" Then
Else
    sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(listb.Value, sh.Range("1:1"), 0)), order1:=xlAscending, Header:=xlYes
End If
End Sub
   
Sub SortZa(nameSh As String, ByVal listb As MSForms.ComboBox)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.sheets(nameSh)
If listb.Value = "ALL" Or listb.Value = "" Then
Else
    sh.UsedRange.Sort key1:=sh.Cells(1, Application.WorksheetFunction.Match(listb.Value, sh.Range("1:1"), 0)), order1:=xlDescending, Header:=xlYes
End If
End Sub


'---------------------UPDATE------------------------

Sub UpdateValue(ByVal List As MSForms.ListBox, sheetSource As String, keysheet As String, ByVal keyId As MSForms.TextBox)
'UP DATE DATAS IN LIST GIAY
If List.ListIndex < 0 Then
    MsgBox "Vui long chon 1 ban ghi", vbInformation
Else
    Dim sheett As Worksheet
    Set sheett = ThisWorkbook.sheets(sheetSource)
    Dim dt As Long
    dt = Application.WorksheetFunction.Match(CInt(keyId.Value), sheett.Range("A:A"), 0)
    Row = Application.WorksheetFunction.CountA(sheett.Range("A:A"))
    For count = 1 To Application.WorksheetFunction.CountA(sheett.Rows(1))
    sheett.Cells(dt, count).Value = UserForm1.Controls(keysheet & count).Value
    Next count

    If keysheet = "LG" Then
        Call UserForm1.ShowlistLG
        Call clearDataLG
     ElseIf keysheet = "G" Then
        Call UserForm1.ShowlistG
        Call ClearDataG
     ElseIf keysheet = "K" Then
        Call UserForm1.ShowlistK
        Call clearDataKH
     ElseIf keysheet = "LO" Then
        Call UserForm1.ShowlistLO
        Call clearDataLO
    End If

Call UserForm1.SetVL
    MsgBox "Du lieu da update"
End If
End Sub


'--------------------CHANGE COLOR-------------------------

Sub Change(ByVal MaGiay As MSForms.ComboBox, ByVal valGiay As MSForms.ComboBox)
If Application.WorksheetFunction.CountIf(ThisWorkbook.sheets("LIST_GIAY").Range("B:B"), (Trim(MaGiay.Value) & Trim(valGiay.Value))) = 0 Then
    MaGiay.BackColor = &HFF&
    valGiay.BackColor = &HFF&
' DOI MAU GIAY
Else
    MaGiay.BackColor = &H80000005
    valGiay.BackColor = &H80000005
End If
End Sub

'======================== FORM ===============================



Private Sub ExportSP_Click()
'EXPORT FILE KHACH HANG
    Dim nS As export
    Set nS = New export
    nS.NewShPreview "SHOW_SP"
End Sub

Private Sub LG9_Change()
    Me.LG9.Value = CStr(Me.LG9.Value)
End Sub

Private Sub MultiPage1_Change()

End Sub

Private Sub P4_AfterUpdate()
    Call P4_Change
End Sub

Private Sub P4_Change()
'SET FOR VALUE
If Application.WorksheetFunction.CountIf(sheets("SP").Range("D:D"), Me.P4.Value) <> 0 And Me.P4.Value <> "" Then
    Me.GROUP.Value = Application.WorksheetFunction.CountIf(sheets("SP").Range("D:D"), Me.P4.Value) + 1
Else
    Me.GROUP.Value = 1
End If
End Sub

Private Sub P57_Change()
    Me.P57.Value = Format(Me.P57.Value, "dd,mm,yyyy")
End Sub

Private Sub P7_Change()
    Me.P7.Value = Format(Me.P7.Value, "#,##")
End Sub

Private Sub P9_Change()
    Me.P9.Value = Format(Me.P7.Value, "#,##")
End Sub

'=======Show ra Form=====
Private Sub ShowForm_Click()
    Dim sheet As Worksheet
    Set sheet = Application.Worksheets("FORM")
    Dim i As Integer
    For i = 2 To 56
        sheet.Range("GP." & i).Value = Me.Controls("P" & i).Value
    Next i
    sheet.Range("GP.57").Value = Me.TextBox54.Value
    sheet.Range("GP.58").Value = Me.TextBox55.Value
    UserForm1.Hide
    sheet.PrintPreview
End Sub


'=======mouse scroll
Private Sub L1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.L1
End Sub
Private Sub L2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.L2
End Sub
Private Sub L3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.L3
End Sub
Private Sub L4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.L4
End Sub
Private Sub L5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.L5
End Sub


'======SORT======
Private Sub az_Click()
    Dim az As SortData
    Set az = New SortData
az.SortAz "SHOW_SP", S1
End Sub

Private Sub za_Click()
    Dim az As SortData
    Set az = New SortData
    az.SortZa "SHOW_SP", S1
End Sub


'show giay


Private Sub azLG_Click()
    Dim azLG As SortData
    Set azLG = New SortData
    azLG.SortAz "SHOW_LIST_GIAY", S3
End Sub

Private Sub zaLG_Click()
    Dim azLG As SortData
    Set azLG = New SortData
    azLG.SortZa "SHOW_LIST_GIAY", S3
End Sub



Private Sub ClearDataG1_Click()
    Call ClearDataG
End Sub

Private Sub ClearFormKH_Click()
    Call clearDataKH
End Sub

Private Sub ClearFormLG_Click()
'CLEAR DATA LIST PAPER
    Call clearDataLG
End Sub

Private Sub ClearFormLO_Click()
'CLEAR DATA FORM LOAI SONG
    Call clearDataLO
End Sub

Private Sub ClearFormSp_Click()
'CLEAR DATA FORM SP
    Call clearDataSp
End Sub


'=====================Export new File=============
Private Sub ExportKH_Click()
'EXPORT FILE KHACH HANG
    Dim n As export
    Set n = New export
    n.NewShPreview "KH"
End Sub
Private Sub ExportLG_Click()
'EXPORT FILE LIST GIAY
    Dim n As export
    Set n = New export
    n.NewShPreview "SHOW_LIST_GIAY"
End Sub

Private Sub ExportG_Click()
'EXPORT FILE LOAI GIAY
    Dim n As export
    Set n = New export
    n.NewShPreview "LOAI_GIAY"
End Sub

Private Sub ExportLO_Click()
'EXPORT FILE LOAI SONG
    Dim n As export
    Set n = New export
    n.NewShPreview "LOAI_SONG"
End Sub

'======================Create======================

Private Sub CreateLO_Click()
'CREATE DU LIEU LIST LOAI SONG
    Dim LO As create
    Set LO = New create
    LO.CreateData "LOAI_SONG", LO2, "B", "LO"
'Call clearDataLO
    Call ShowlistLO
End Sub

Private Sub CreateLG_Click()
'CREATE DU LIEU LIST GIAY
    Dim LG As create
    Set LG = New create
    LG.CreateNoKey "LIST_GIAY", LG1, "A", "LG"
'all clearDataLG
    Call ShowlistLG
End Sub

Private Sub CreateG_Click()
'CREATE DU LIEU GIAY
    Dim G As create
    Set G = New create
    G.CreateData "LOAI_GIAY", G3, "C", "G"
Call ClearDataG
Call ShowlistG
Call ShowlistG
Call SetVL
End Sub

Private Sub CreateKH_Click()
'CREATE DU LIEU KH
    Dim KH As create
    Set KH = New create
    KH.CreateData "KH", K2, "B", "K"
'Call clearDataKH
Call ShowlistK
End Sub

Private Sub CreateP_Click()
'CREATE DU LIEU SP
    Dim SP As create
    Set SP = New create
'*********************
    SP.CreateDataKeyId "SP", P10, "J", "P"
'Call clearDataSp
'Call ShowlistB
End Sub





'===========================DELETE ROWS=======================

Private Sub DeleteDataSp_Click()
'DELETE ROW
    Dim DL As DeleteData
    Set DL = New DeleteData
    DL.delete "SP", L1
    Call clearDataSp
    Call ShowlistB
End Sub

Private Sub DeleteDataKH_Click()
    Dim DL As DeleteData
    Set DL = New DeleteData
    DL.delete "KH", L2
Call clearDataKH
Call ShowlistK
End Sub

Private Sub DeleteDataLO_Click()
    Dim DL As DeleteData
    Set DL = New DeleteData
    DL.delete "LOAI_SONG", L5
Call clearDataLO
Call ShowlistLO
End Sub

Private Sub DeleteDataG_Click()
    Dim DL As DeleteData
    Set DL = New DeleteData
    DL.delete "LOAI_GIAY", L3
Call ClearDataG
Call ShowlistG
End Sub



'==========================SHOW DATA=======================

Private Sub L1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'TRA DU LIEU TU DATA RA FORM
    Dim na As Worksheet
    Set na = ThisWorkbook.sheets("SP")
For i = 34 To 42
    step = 2
    Me.Controls("P" & i).Style = 0
    Me.Controls("P" & 3).Style = 0
Next i
    For count = 1 To Application.WorksheetFunction.CountA(na.Rows(1))
    Me.Controls("P" & count).Value = Me.L1.List(Me.L1.ListIndex, count - 1)
    Next count
End Sub

Private Sub L2_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'TRA DU LIEU TU DATA RA FORM SHEET KH
    Dim nameS As Worksheet
    Set nameS = ThisWorkbook.sheets("KH")
For count = 1 To Application.WorksheetFunction.CountA(nameS.Rows(1))
    Me.Controls("K" & count).Value = Me.L2.List(Me.L2.ListIndex, count - 1)
Next count
End Sub

Private Sub L3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'TRA DU LIEU TU DATA RA FORM SHEET LOAI_GIAY
    Dim nam As Worksheet
    Set nam = ThisWorkbook.sheets("LOAI_GIAY")
For count = 1 To Application.WorksheetFunction.CountA(nam.Rows(1))
    Me.Controls("G" & count).Value = Me.L3.List(Me.L3.ListIndex, count - 1)
Next count
End Sub

Private Sub L4_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'TRA DU LIEU TU DATA RA FORM SHEET LOAI_GIAY
    Dim n As Worksheet
    Set n = ThisWorkbook.sheets("LIST_GIAY")
For count = 1 To Application.WorksheetFunction.CountA(n.Rows(1))
    Me.Controls("LG" & count).Value = Me.L4.List(Me.L4.ListIndex, count - 1)
Next count
End Sub

Private Sub L5_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'TRA DU LIEU TU DATA RA FORM SHEET LOAI_GIAY
    Dim nameShP As Worksheet
    Set nameShP = ThisWorkbook.sheets("LOAI_SONG")
For count = 1 To Application.WorksheetFunction.CountA(nameShP.Rows(1))
    Me.Controls("LO" & count).Value = Me.L5.List(Me.L5.ListIndex, count - 1)
Next count
End Sub


Private Sub S2_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'TIM KIEM DU LIEU
 If KeyCode = 13 And S1.Value <> "" Then
        Call ShowlistB
    End If
End Sub

Private Sub S4_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'TIM KIEM DU LIEU
 If KeyCode = 13 And S3.Value <> "" Then
        Call ShowlistLG
    End If
End Sub

'========================CHANGE du lieu===================

Private Sub LO2_AfterUpdate()
    Call LO2_Change
End Sub
Private Sub LO2_Change()
If Len(Trim(Me.LO2.Value)) > 1 Then
    Me.LO6.Visible = True
    Me.LO7.Visible = True
    Me.SO8.Visible = True
    Me.SO9.Visible = True
Else
    Me.LO6.Visible = False
    Me.LO7.Visible = False
    Me.LO6.Value = ""
    Me.LO7.Value = ""
    Me.SO8.Visible = False
    Me.SO9.Visible = False
'TRA VE ID

End If
End Sub

'THAY DOI LOAI SONG
 Private Sub P11_change()
    Call P11_AfterUpdate
 End Sub
Private Sub P11_AfterUpdate()
'LOAI SONG SP
If Len(Trim(Me.P11.Value)) > 1 Then
    For i = 1 To 4
        Me.Controls("labe" & i).Visible = True
    Next i
    For i = 38 To 41
        Me.Controls("P" & i).Visible = True
    Next i
        Me.TextBox59.Visible = True
        Me.TextBox60.Visible = True
'THAY DOI LOAI SONG BEN MEDIUM
        Me.Labe3.Caption = "sóng " & Application.WorksheetFunction.VLookup(P11.Value, Worksheets("LOAI_SONG").Range("B:G"), 5, 0)
        Me.Labe4.Caption = "sóng " & Application.WorksheetFunction.VLookup(P11.Value, Worksheets("LOAI_SONG").Range("B:G"), 6, 0)
        Me.Labe10.Caption = "Medium 1"
'CLOSE
Else
        Me.TextBox59.Value = ""
        Me.TextBox60.Value = ""
        Me.TextBox59.Visible = False
        Me.TextBox60.Visible = False
    For ii = 38 To 41
        Me.Controls("P" & ii).Style = 0
        Me.Controls("P" & ii).Visible = False
        Me.Controls("P" & ii).Value = ""
    Next ii
    For i = 1 To 4
        Me.Controls("labe" & i).Visible = False
    Next i
        Me.Labe10.Caption = "Medium"
End If


End Sub

'======CHANGE GIAY======

Private Sub P35_Click()
    Call P35_Change
End Sub
Private Sub P35_AfterUpdate()
    Call P35_Change
End Sub
Private Sub P35_Change()
    Dim P35ch As changeData
    Set P35ch = New changeData
    P35ch.Change P34, P35
Call ShowNote
End Sub

Private Sub P37_AfterUpdate()
Call P37_Change
End Sub

Private Sub P37_Change()
    Dim P37ch As changeData
    Set P37ch = New changeData
    P37ch.Change P36, P37
Call ShowNote
End Sub

Private Sub P39_AfterUpdate()
Call P39_Change
End Sub

Private Sub P39_Change()
    Dim P39ch As changeData
    Set P39ch = New changeData
    P39ch.Change P38, P39
Call ShowNote
End Sub

Private Sub P41_AfterUpdate()
Call P41_Change
End Sub

Private Sub P41_Change()
    Dim P41ch As changeData
    Set P41ch = New changeData
    P41ch.Change P40, P41
Call ShowNote
End Sub

Private Sub P43_AfterUpdate()
Call P43_Change
End Sub

Private Sub P43_Change()
    Dim h As Worksheet
    Set h = ThisWorkbook.sheets("LOAI_GIAY")
    Dim P43ch As changeData
    Set P43ch = New changeData
    P43ch.Change P42, P43
Call ShowNote
 'SET TEN CHO TEN NHOM
Call SetNameGroup
End Sub
'======================
Private Sub LG5_Change()
Call LG3_Change
End Sub
Private Sub LG3_Change()
'tim kiem ten ncc giay
     Me.LG2.Value = Trim(Me.LG3.Value) & Trim(Me.LG5.Value)
If Me.LG3.Value <> "" Then 'set value for LG4
    Me.LG4.Value = WorksheetFunction.IfError(Application.WorksheetFunction.VLookup(Me.LG3.Value, Worksheets("LOAI_GIAY").Range("C:D"), 2, 0), "")
End If
End Sub


Private Sub P3_Change()
'tim kiem id khach hang
If Me.P3.Value <> "" Then
    If Application.WorksheetFunction.CountIf(sheets("KH").Range("B:B"), Me.P3.Value) = 0 Then
        MsgBox "MA KHACH HANG SAI, NHAP LAI", vbInformation
    Else
        Me.P2.Value = Application.WorksheetFunction.VLookup(Me.P3.Value, sheets("KH").Range("B:E"), 2, 0)
        TextBox54.Value = Application.WorksheetFunction.VLookup(Me.P3.Value, sheets("KH").Range("B:E"), 3, 0)
        TextBox55.Value = Application.WorksheetFunction.VLookup(Me.P3.Value, sheets("KH").Range("B:E"), 4, 0)
        TextBox56.Value = Application.WorksheetFunction.VLookup(Me.P3.Value, sheets("KH").Range("B:F"), 5, 0)
    If Application.WorksheetFunction.VLookup(Me.TextBox56.Value, sheets("SOURCE").Range("T:U"), 2, 0) = 0 Then
        TextBox56.BackColor = &HFF&
    Else
        TextBox56.BackColor = &H80000005
    End If
    End If
End If
End Sub

'SCROLL MOUSE MOVE
Private Sub P3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P3
End Sub
Private Sub LG3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.LG3
End Sub
Private Sub G4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.G4
End Sub
Private Sub G5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.G5
End Sub
Private Sub P34_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P34
End Sub
Private Sub P35_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P35
End Sub
Private Sub P36_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P36
End Sub
Private Sub P37_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P37
End Sub
Private Sub P38_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P38
End Sub
Private Sub P39_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P39
End Sub
Private Sub P40_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P40
End Sub
Private Sub P41_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P41
End Sub
Private Sub P42_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    HookListBoxScroll Me, Me.P42
End Sub

Private Sub P34_Change()
        Dim Class As SetValue
        Set Class = New SetValue
' SL GIAY
        Class.NewValueRange "LIST_GIAY", P35, "C", P34, "E", "R", "M"
'SET TEN CHO TEN NHOM
Call SetNameGroup
  'TIM TRANG THAI
    If Me.P34.Value <> "" Then
        If Application.WorksheetFunction.CountIf(sheets("LOAI_GIAY").Range("C:C"), Me.P34.Value) = 0 Then
            MsgBox "MA GIAY NHAP SAI"
         Else
            Me.TextBox57.Value = Application.WorksheetFunction.VLookup(P34.Value, sheets("LOAI_GIAY").Range("C:I"), 7, 0)
            If Application.WorksheetFunction.VLookup(Me.TextBox57.Value, sheets("SOURCE").Range("R:S"), 2, 0) = 0 Then
                Me.TextBox57.BackColor = &HFF&
' DOI MAU GIAY
            Else
                Me.TextBox57.BackColor = &H80000005
            End If
        End If
    End If
End Sub

Private Sub P36_Change()
        Dim Class As SetValue
        Set Class = New SetValue
' SL GIAY
        Class.NewValueRange "LIST_GIAY", P37, "C", P36, "E", "R", "M"
'SET TEN CHO TEN NHOM
Call SetNameGroup
'TIM TRANG THAI
    If Me.P36.Value <> "" Then
        If Application.WorksheetFunction.CountIf(sheets("LOAI_GIAY").Range("C:C"), Me.P36.Value) = 0 Then
            MsgBox "MA GIAY NHAP SAI"
        Else
            Me.TextBox58.Value = Application.WorksheetFunction.VLookup(P36.Value, sheets("LOAI_GIAY").Range("C:I"), 7, 0)
            If Application.WorksheetFunction.VLookup(Me.TextBox58.Value, sheets("SOURCE").Range("R:S"), 2, 0) = 0 Then
                Me.TextBox58.BackColor = &HFF&
' DOI MAU GIAY
            Else
                Me.TextBox58.BackColor = &H80000005
            End If
        End If
    End If
End Sub
Private Sub P38_Change()
        Dim Class As SetValue
        Set Class = New SetValue
' SL GIAY
        Class.NewValueRange "LIST_GIAY", P39, "C", P38, "E", "R", "M"
 'SET TEN CHO TEN NHOM
Call SetNameGroup
If P38.Value <> "" Then
    If Application.WorksheetFunction.CountIf(sheets("LOAI_GIAY").Range("C:C"), Me.P38.Value) = 0 Then
        MsgBox "MA GIAY NHAP SAI"
    Else
    'TIM TRANG THAI
        Me.TextBox59.Value = Application.WorksheetFunction.VLookup(P38.Value, sheets("LOAI_GIAY").Range("C:I"), 7, 0)
        If Application.WorksheetFunction.VLookup(Me.TextBox59.Value, sheets("SOURCE").Range("R:S"), 2, 0) = 0 Then
            Me.TextBox59.BackColor = &HFF&
' DOI MAU GIAY
        Else
            Me.TextBox59.BackColor = &H80000005
        End If
    End If
End If
End Sub
Private Sub P42_Change()
    Dim Class As SetValue
    Set Class = New SetValue
  ' SL GIAY
        Class.NewValueRange "LIST_GIAY", P43, "C", P42, "E", "R", "M"
'SET TEN CHO TEN NHOM
Call SetNameGroup
    If P42.Value <> "" Then
      'TIM TRANG THAI
        If Application.WorksheetFunction.CountIf(sheets("LOAI_GIAY").Range("C:C"), Me.P42.Value) = 0 Then
             MsgBox "MA GIAY NHAP SAI"
        Else
            Me.TextBox61.Value = Application.WorksheetFunction.VLookup(P42.Value, sheets("LOAI_GIAY").Range("C:I"), 7, 0)
            If Application.WorksheetFunction.VLookup(Me.TextBox61.Value, sheets("SOURCE").Range("R:S"), 2, 0) = 0 Then
                Me.TextBox61.BackColor = &HFF&
' DOI MAU GIAY
            Else
                Me.TextBox61.BackColor = &H80000005
            End If
        End If
     End If
End Sub

Private Sub P40_Change()
     Dim Class As SetValue
    Set Class = New SetValue
  ' SL GIAY
    Class.NewValueRange "LIST_GIAY", P41, "C", P40, "E", "R", "M"
'SET TEN CHO TEN NHOM
Call SetNameGroup
  If P40.Value <> "" Then
  'TIM TRANG THAI
    If Application.WorksheetFunction.CountIf(sheets("LOAI_GIAY").Range("C:C"), Me.P40.Value) = 0 Then
        MsgBox "MA GIAY NHAP SAI"
    Else
        Me.TextBox60.Value = Application.WorksheetFunction.VLookup(P40.Value, sheets("LOAI_GIAY").Range("C:I"), 7, 0)
        If Application.WorksheetFunction.VLookup(Me.TextBox60.Value, sheets("SOURCE").Range("R:S"), 2, 0) = 0 Then
         Me.TextBox60.BackColor = &HFF&
' DOI MAU GIAY
         Else
         Me.TextBox60.BackColor = &H80000005
         End If
    End If
  End If
End Sub


Private Sub P46_Change()
'IN AN
If Me.P46.Value = "NO" Or Me.P46.Value = "-" Then
    Me.P47.Visible = False
    Me.P48.Visible = False
    Me.P47.Value = ""
    Me.P48.Value = ""
    Me.Labe38.Visible = False
    Me.Labe39.Visible = False
Else
    Me.P47.Visible = True
    Me.P48.Visible = True
    Me.Labe38.Visible = True
    Me.Labe39.Visible = True
End If
End Sub

'TEN KHACH HANG
Private Sub K2_Change()
    Me.K3.Value = Me.K1.Value
End Sub

'=========================UPDATE=========================

Private Sub UpdateKH_Click()
'UP DATE DATAS IN KH
    Dim UpPPPP As Update
    Set UpPPPP = New Update
    UpPPPP.UpdateValue L2, "KH", "K", K1
End Sub

Private Sub UpdateLG_Click()
'UPDATE LIST GIAY
    Dim UpPPP As Update
    Set UpPPP = New Update
    UpPPP.UpdateValue L4, "LIST_GIAY", "LG", LG1
End Sub

Private Sub UpdateG_Click()
'UP DATE DATA IN GIAY
    Dim UpPP As Update
    Set UpPP = New Update
    UpPP.UpdateValue L3, "LOAI_GIAY", "G", G1
End Sub

Private Sub UpdateLO_Click()
'LOAI SONG
    Dim UpP As Update
    Set UpP = New Update
    UpP.UpdateValue L5, "LOAI_SONG", "LO", LO1
End Sub

Private Sub UpdateSp_Click()
'UP DATE DATAS IN SP
If L1.ListIndex < 0 Then
    MsgBox "Vui long chon 1 ban ghi", vbInformation
Else
    Dim sh As Worksheet
    Set sh = ThisWorkbook.sheets("SP")
    Dim d As Long
    d = Application.WorksheetFunction.Match(CInt(Me.P1.Value), sh.Range("A:A"), 0)
    Row = Application.WorksheetFunction.CountA(sh.Range("A:A"))
    For count = 1 To Application.WorksheetFunction.CountA(sh.Rows(1))
        If Not count = 57 Then
            sh.Cells(d, count).Value = Me.Controls("P" & count).Value
        End If
    Next count
    sh.Cells(d, 58).Value = Now
Call clearDataSp
Call ShowlistB
Call SetVL
    MsgBox "Du lieu da update"
End If
End Sub

'=====================SHOW LIST=================

Sub ShowlistB()
'SHOW LIST BOX SAN PHAM
  Dim Showlist1 As ShowList
  Set Showlist1 = New ShowList
    Showlist1.showListFilter "SP", "SHOW_SP", S1, S2, L1
With Me.L1
.ColumnWidths = "30,55,80,120,360,80,70,60,60"
End With
End Sub

Sub ShowlistK()
'SHOW LIST BOX KH
  Dim Showlist3 As ShowList
  Set Showlist3 = New ShowList
     Showlist3.ShowListBox "KH", L2
 With Me.L2
    .ColumnWidths = "30,120,30,290"
End With
End Sub

Sub ShowlistG()
'SHOW LIST BOX KH
  Dim Showlist2 As ShowList
  Set Showlist2 = New ShowList
 Showlist2.ShowListBox "LOAI_GIAY", L3
 With Me.L3
.ColumnWidths = "30"
End With
End Sub

Sub ShowlistLG()
'SHOW LIST BOX KH
    Dim Showlist1 As ShowList
  Set Showlist1 = New ShowList
 Showlist1.showListFilter "LIST_GIAY", "SHOW_LIST_GIAY", S3, S4, L4
 With Me.L4
.ColumnWidths = "30"
End With
End Sub

Sub ShowlistLO()
'SHOW LIST BOX KH
  Dim Showlist2 As ShowList
  Set Showlist2 = New ShowList
 Showlist2.ShowListBox "LOAI_SONG", L5
 With Me.L5
.ColumnWidths = "30"
End With
End Sub

'=========================SET VALUE=======================

Private Sub UserForm_Initialize()
'SET VALUE CHO CAC GIA TRI
  Dim Class As SetValue
  Set Class = New SetValue
    Me.S1.Value = "ALL"
    Me.S3.Value = "ALL"
    Me.P57.Value = Format(Date, "DD-MM-YYYY")
    Call SetVL
    Call ShowlistB
    Call ShowlistK
    Call ShowlistG
    Call ShowlistLG
    Call ShowlistLO
    Me.Width = 1080
    Me.Height = 600
End Sub

Sub SetVL()
'SET VALUE CHO CAC COMBOBOX
  Dim Class As SetValue
  Set Class = New SetValue
'KHACH HANG
Class.NewValueFilter "KH", "B", P3, "F", "SOURCE", "T"
'KH
Class.NewValue "SOURCE", "T", K6
'Tim kiem
Class.NewValue "SOURCE", "A", S1
'TIMKIEM LIST GIAY
Class.NewValue "SOURCE", "Z", S3
' LOAI SP
Class.NewValue "SOURCE", "C", P6
' chu ki
Class.NewValue "SOURCE", "B", P8
'loai song
Class.NewValue "LOAI_SONG", "B", P11
'TlT trên/
Class.NewValue "SOURCE", "D", P15
'Tl trên/
Class.NewValue "SOURCE", "E", P20
'Don vi tinh do buc, do nen
Class.NewValue "SOURCE", "H", P22
 'Don vi tinh do buc, do neN
Class.NewValue "SOURCE", "H", P24
'Don vi tinh Chong tham
Class.NewValue "SOURCE", "I", P26
 'Don vi tinh Chong tham
Class.NewValue "SOURCE", "I", P28
 'Dung thung
Class.NewValue "SOURCE", "G", P32
 'Loai giay
Class.NewValueFilter "LOAI_GIAY", "C", P34, "I", "SOURCE", "R"
 'Loai giay
Class.NewValueFilter "LOAI_GIAY", "C", P36, "I", "SOURCE", "R"
 'Loai giay
Class.NewValueFilter "LOAI_GIAY", "C", P38, "I", "SOURCE", "R"
 'Loai giay
Class.NewValueFilter "LOAI_GIAY", "C", P40, "I", "SOURCE", "R"
 'Loai giay
Class.NewValueFilter "LOAI_GIAY", "C", P42, "I", "SOURCE", "R"
 'In an
Class.NewValue "SOURCE", "L", P46
 'TÊN NCC
Class.NewValue "SOURCE", "Q", G4
'TÊN NCC
Class.NewValue "SOURCE", "R", G9
 'TÊN NCC
Class.NewValue "LOAI_GIAY", "C", LG3
'TÊN XEP THUNG
Class.NewValue "SOURCE", "F", P31
'TR?NG THÁI GI?Y
Class.NewValue "SOURCE", "V", LG13
'GIAO NHAN
Class.NewValue "SOURCE", "K", P49

'DAC DIEM GIAO NHAN
Class.NewValue "SOURCE", "J", P50

'DAC DIEM GIAO NHAN
Class.NewValue "SOURCE", "X", G5
End Sub
