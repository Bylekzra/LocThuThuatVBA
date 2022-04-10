Attribute VB_Name = "ThuThuat0YHCT0BSHoang"
Sub ALL_XuLySoLieu()    'Ctrl+Shift+D
Attribute ALL_XuLySoLieu.VB_ProcData.VB_Invoke_Func = "D\n14"
Dim i As Variant
Dim str As String
Dim strx As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "So Phau thuat" Then
            strx = str
           Call XuLySoLieu_680(strx)
        End If
        If str = "DuLieu" Then
            strx = str
           Call XuLySoLieu_DuLieu(strx)
        End If
Next
str = ""
strx = ""

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub XuLySoLieu_680(strx As String)
Attribute XuLySoLieu_680.VB_ProcData.VB_Invoke_Func = " \n14"
Dim i As Variant
Dim j As Variant
Dim m As Variant
Dim str As String
Dim Str2 As String
Dim ww As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'On Error Resume Next
For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "Cacche1" Then
           Str2 = "H"
        End If
Next
On Error Resume Next
If Str2 = "H" Then
 Sheets("PivotH").Delete
 Sheets("Cacche1").Delete
 Sheets("So Phau thuat").Delete
 Sheets("So Phau thuat (2)").Name = "So Phau thuat"
 Else
 Sheets(1).Delete
End If
str = ""
Str2 = ""
'Coppy sheet
    Sheets("So Phau thuat").Copy Before:=Sheets(1)
    Sheets("So Phau thuat").Move Before:=Sheets(1)
   
Sheets("So Phau thuat").Select

Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight
Columns(14).Insert Shift:=xlToRight

Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight

Columns(9).Insert Shift:=xlToRight
Columns(9).Insert Shift:=xlToRight
Columns(9).Insert Shift:=xlToRight

'Doi ten cells
    Cells(6, 3) = "BoRn"
    Cells(6, 5) = "DTbn"
    Cells(6, 9) = "TenDv"
    Cells(6, 10) = "GiaDv"
    Cells(6, 11) = "TimeBH"
    Cells(6, 13) = "ngaygio_lam"
    Cells(6, 15) = "NgayStart"
    Cells(6, 16) = "GioPhut"
    Cells(6, 17).FormulaR1C1 = "Time" & Chr(10) & "Vao-Ra"
    Cells(6, 18).FormulaR1C1 = "T7" & Chr(10) & "CN"
    Cells(6, 19) = "00Sec"
    Cells(6, 20) = "00Sec2"
    Cells(6, 21).FormulaR1C1 = "Sum" & Chr(10) & "TimeBH"
    Cells(6, 22) = "LoaiDv"
    Cells(6, 23) = "LoaiNv"
    Cells(6, 24) = "TenNv"
    Cells(6, 27) = "1BN-nNV"
    Cells(6, 28) = "1Nv-nBN"
    Cells(6, 29) = "NoTruc"
    Cells(6, 30) = "NgTruc"
    Cells(6, 31) = "ChâmTruc"
    Cells(6, 32) = "FixTT"  'Chua Biet Lam Gi
    Cells(6, 33) = "CCHN"
    Cells(6, 34) = "."
    Cells(6, 35) = "Cacche"
    'Cells(6, 36) = "."
    'Cells(6, 37) = "."
'Het doi ten cells
    Cells(1, 4) = "4"
    Cells(1, 10) = "10"
    Cells(1, 14) = "14"
    Cells(1, 18) = "18"
    Cells(1, 24) = "24"
    Cells(1, 30) = "30"
    Cells(1, 35) = "35"
'Can vi tri cot

Range("K:K").NumberFormat = "mm:ss"
Range("Q:Q").NumberFormat = "[h]:mm:ss;@" 'Format gio phut giay
Range("P:P").NumberFormat = "h:mm;@"
Range("O:O").NumberFormat = "m/d/yyyy"
Range("S:T").NumberFormat = "hh:mm dd/mm/yyyy"
Range("U:U").NumberFormat = "hh:mm"
Columns("V:W").NumberFormat = "General"
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").ColumnWidth = 14
Columns("C:C").ColumnWidth = 3.5
Columns("E:E").ColumnWidth = 0.5
Columns("F:H").ColumnWidth = 0.5
Columns("I:I").ColumnWidth = 7.2
Columns("J:J").ColumnWidth = 4.5
Columns("K:K").ColumnWidth = 2
Columns("L:L").ColumnWidth = 0.5
Columns("M:N").EntireColumn.AutoFit
Columns("O:O").ColumnWidth = 8.2
Columns("P:P").ColumnWidth = 4.1
Columns("Q:Q").ColumnWidth = 5.5
Columns("R:R").ColumnWidth = 2.3
Columns("S:T").ColumnWidth = 0.5
Columns("U:U").ColumnWidth = 3.8
Columns("V:W").ColumnWidth = 4
Columns("X:X").ColumnWidth = 12
Columns("Y:Y").ColumnWidth = 0.5
Columns("Z:Z").ColumnWidth = 1
Columns("AJ:AK").ColumnWidth = 4.5
Columns("AA:AI").EntireColumn.AutoFit
Rows(6).AutoFilter

'Danh dau o can xem
Cells(5, 17) = ">20'"
Cells(5, 21) = ">8:00"
Cells(5, 22) = "T"
Cells(5, 23) = "T"
Cells(5, 24) = "Blank"
Cells(5, 27) = "#1"
Cells(5, 28) = "#1"
Cells(5, 29) = "Value"
    
'Tao Cacche1
Call Cacche1
ww = Sheets("Cacche1").Cells(5, 17)

Sheets("So Phau thuat").Select

Cells(1, 2) = ""
i = 7
While Cells(i, 8) <> ""
    'Xoa doi tuong <> BHYT
    str = Cells(i, 37)
        If str <> "" Then
        'If str = "BHYT" Then
        'Doi ten thu thuat sang 21 ky tu
            Cells(i, 12) = Left(Cells(i, 8), 21)
        'Chuyen ten nguoi lam ve dung cot
            'Cells(i, 17) = Left(Cells(i, 16), 50) & Left(Cells(i, 15), 50) & Left(Cells(i, 14), 50)
        'Format thoi gian lam
            Cells(i, 15).FormulaR1C1 = "=DATE(YEAR(RC[-2]),MONTH(RC[-2]),DAY(RC[-2]))"
            'Cells(i, 18) = Left(Cells(i, 11), 10)
        'Time vào - time ra
            Cells(i, 17) = Cells(i, 14) - Cells(i, 13)
            Cells(i, 9).FormulaR1C1 = "=bo_dau_tieng_viet(RC[3])"
            Cells(i, 24).FormulaR1C1 = "=bo_dau_tieng_viet(RC[1])"
        'GiaDv
            Cells(i, 10).FormulaR1C1 = "=VLOOKUP(RC[-1], Cacche1!R1C2:R" & ww - 2 & "C8,7,0)"
        'Loai Dv Loai Nv
            Cells(i, 22).FormulaR1C1 = "=VLOOKUP(RC[-13], Cacche1!R1C2:R" & ww - 2 & "C3,2,0)"
            Cells(i, 23).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[1], Cacche1!R1C5:R6C6,2,0),""Ca2"")"
        'Add TimeBH
            Cells(i, 11).FormulaR1C1 = "=VLOOKUP(RC[-2], Cacche1!R1C2:R" & ww - 2 & "C7,6,0)"
        'Loc T7/CN
            Cells(i, 18).FormulaR1C1 = "=IF(WEEKDAY(RC[-5])=1,""CN"",IF(WEEKDAY(RC[-5])=7,""T7"",IF(WEEKDAY(RC[-5])=6,""T6"", IF(WEEKDAY(RC[-5])=5,""T5"", IF(WEEKDAY(RC[-5])=4,""T4"", IF(WEEKDAY(RC[-5])=3,""T3"", IF(WEEKDAY(RC[-5])=2,""T2"")))))))"
        'Doi ten thu thuat cho pivot de nhin
            'Cells(i, 9).FormulaR1C1 = "=VLOOKUP(bo_dau_tieng_viet(Left(Cells(RC[-1], 21)), Cacche1!R1C2:R15C4,3,0)"
        'Chuyen Gio Phut bat dau lam
            Cells(i, 16).FormulaR1C1 = "=TIME(HOUR(RC[-2]),MINUTE(RC[-2]),SECOND(RC[-2]))"
        '00Sec
            Cells(i, 19).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
        '00Sec2
            Cells(i, 20).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
            
            i = i + 1
            m = i
        Else
                Rows(i).Delete
        End If
    'Het xoa doi tuong <> BHYT
Wend
i = 0
str = ""

Call PasteValueH(strx, m)

'
Call CCHN_Error(m)

'Doi ten thu thuat cho pivot de nhin
Call ShortNameTT(strx, m, ww)

'Loc trung gio2
Application.Calculation = xlCalculationAutomatic 'Cái này de tinh Automatic moi coppy/paste duoc
       Cells(7, 27).Value = _
            "=COUNTIFS($B$7:$B$" & m - 1 & ",B7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7,$Y$7:$Y$" & m - 1 & ",""<>""&Y7)+COUNTIFS($B$7:$B$" & m - 1 & ",B7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7)" 'Loc cot Phu second = 00
       Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
       Cells(7, 28).Value = _
            "=COUNTIFS($Y$7:$Y$" & m - 1 & ",Y7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7,$B$7:$B$" & m - 1 & ",""<>""&B7)+COUNTIFS($Y$7:$Y$" & m - 1 & ",Y7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7)" 'Loc cot Phu second = 00
       Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
       'Sum TimeBH
       Cells(7, 21).Value = "=SUMIFS($K$7:$K$" & m - 1 & ",$Y$7:$Y$" & m - 1 & ",Y7,$O$7:$O$" & m - 1 & ",O7)"
       Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
       'Dem loc trung gio
       'Cells(7, 33).Value = "=AA7+AB7"
       
Application.Calculation = xlCalculationManual 'Het lenh Cái này de tinh Automatic moi coppy/paste duoc
    
    Range("AA7:AB" & m - 1).Copy 'Ctrl+Shift+D thì bo lenh copy 2lenh
    Range("AA7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("U7:U" & m - 1).Copy
    Range("U7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Hêt Loc trung gio2

'Tich nham nguoi khong truc
Call NoTruc(strx, m, ww)

'Color
Call CoLor(m)

'So Luong BN theo ngay
Call BN_Date(strx, m, ww)

'PivotH
    Sheets.Add.Name = "PivotH"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "So Phau thuat!R6C1:R150000C37", Version:=6).CreatePivotTable TableDestination _
        :="PivotH!R2C1", TableName:="PivotTableH", DefaultVersion:=6
    Sheets("PivotH").Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTableH")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTableH").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTableH").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("NgayStart")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("nguoi_thuc_hien")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("TenDv")
        .Orientation = xlColumnField  'lân 1
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTableH").AddDataField ActiveSheet.PivotTables( _
        "PivotTableH").PivotFields("TenDv"), "Count of TenDv", xlCount
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("TenDv")
        .Orientation = xlColumnField 'lân 2
        .Position = 1
    End With
    'With ActiveSheet.PivotTables("PivotTableH").PivotFields("LoaiDv")
    '    .Orientation = xlColumnField
    '    .Position = 1
    'End With
'PivotH TimeBH
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "So Phau thuat!R6C1:R150000C37", Version:=6).CreatePivotTable TableDestination _
        :="PivotH!R30C1", TableName:="PivotTableHH", DefaultVersion:=6
    Sheets("PivotH").Cells(30, 1).Select
    With ActiveSheet.PivotTables("PivotTableHH")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTableHH").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTableHH").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTableHH").PivotFields("NgayStart")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableHH").PivotFields("nguoi_thuc_hien")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTableHH").AddDataField ActiveSheet.PivotTables( _
        "PivotTableHH").PivotFields("TimeBH"), "Sum of TimeBH", xlSum 'Lân 1

    ActiveSheet.PivotTables("PivotTableHH").AddDataField ActiveSheet.PivotTables( _
        "PivotTableHH").PivotFields("TimeBH"), "Sum of TimeBH", xlSum  'Lân 2
        
    Range("B31:B55").NumberFormat = "[h]:mm;@"
'Hêt PivotH
Cells(1, 5) = "Chu y ô: Blank"

Sheets("So Phau thuat").Select
Cells(1, 1).Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub XuLySoLieu_DuLieu(strx As String) 'Ctrl+Shift+
Dim i As Variant
Dim j As Variant
Dim m As Variant
Dim str As String
Dim Str2 As String
Dim ww As Integer 'Bien so xuong don thu thuat

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

'On Error Resume Next
For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "Cacche1" Then
           Str2 = "H"
        End If
Next
'On Error Resume Next
If Str2 = "H" Then
 Sheets("PivotH").Delete
 Sheets("Cacche1").Delete
 Sheets("DuLieu").Delete
 Sheets("DuLieu (2)").Name = "DuLieu"
End If
str = ""
Str2 = ""

'Coppy sheet
    Sheets("DuLieu").Copy Before:=Sheets(1)
    Sheets("DuLieu").Move Before:=Sheets(1)
   
Sheets("DuLieu").Select

'Xoa data trang
    For i = 1 To 12
    Columns(16).Delete
    Next
    Columns(11).Delete
    Columns(6).Delete
    Columns(5).Delete
    Columns(1).Delete
    Rows(7).Delete
    Rows(6).Delete
    Rows(5).RowHeight = 28
'Het xoa data trang

Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight
Columns(15).Insert Shift:=xlToRight

Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight
Columns(12).Insert Shift:=xlToRight

Columns(10).Insert Shift:=xlToRight
Columns(10).Insert Shift:=xlToRight
Columns(10).Insert Shift:=xlToRight

'Doi ten cells
    Cells(5, 8) = "TenDv"
    Cells(5, 9) = "GiaDv"
    Cells(5, 10) = "TimeBH"
    Cells(5, 11) = "."
    Cells(5, 12) = "."
    Cells(5, 13) = "NgayGioLam"
    Cells(5, 14) = "NgayGioEnd"
    Cells(5, 18) = "TenNv"
    Cells(5, 15) = "NgayStart"
    Cells(5, 16) = "GioPhut"
    Cells(5, 17).FormulaR1C1 = "Time" & Chr(10) & "Vao-Ra"
    Cells(5, 18).FormulaR1C1 = "T7" & Chr(10) & "CN"
    Cells(5, 19) = "00Sec"
    Cells(5, 20) = "00Sec2"
    Cells(5, 21) = "Sum" & Chr(10) & "TimeBH"
    Cells(5, 22) = "LoaiDv"
    Cells(5, 23) = "LoaiNv"
    Cells(5, 24) = "TenNv"
    Cells(5, 25) = "."
    Cells(5, 26) = "."
    Cells(5, 27) = "1BN-nNV"
    Cells(5, 28) = "1NV-nBN"
    Cells(5, 29) = "NoTruc"
    Cells(5, 30) = "NgTruc"
    Cells(5, 31) = "ChamTruc"
    Cells(5, 32) = "FixTT"
    Cells(5, 33) = "CCHN"
    Cells(5, 34) = "."
    Cells(5, 35) = "Cacche"
'Het doi ten cells

'Can vi tri cot
    Cells(1, 4) = "4"
    Cells(1, 11) = "11"
    Cells(1, 14) = "14"
    Cells(1, 18) = "18"
    Cells(1, 20) = "20"
    Cells(1, 22) = "22"
    Cells(1, 30) = "30"
    Cells(1, 35) = "35"
'Can vi tri cot

'Danh dau o can xem
Cells(4, 17) = ">20'"
Cells(4, 21) = ">8:00"
Cells(4, 22) = "T"
Cells(4, 23) = "T"
Cells(4, 24) = "Blank"
Cells(4, 27) = "#1"
Cells(4, 28) = "#1"
Cells(4, 29) = "Value"

Range("J:J").NumberFormat = "mm:ss"
Range("Q:Q").NumberFormat = "[h]:mm:ss;@" 'Format gio phut giay
Range("P:P").NumberFormat = "h:mm;@"
Range("O:O").NumberFormat = "m/d/yyyy"
Range("S:T").NumberFormat = "hh:mm dd/mm/yyyy"
Range("U:U").NumberFormat = "hh:mm"
Columns("A:A").EntireColumn.AutoFit
Columns("B:B").ColumnWidth = 14
Columns("C:C").ColumnWidth = 3.5
Columns("E:E").ColumnWidth = 0.5
Columns("F:G").ColumnWidth = 0.5
Columns("H:H").ColumnWidth = 9
Columns("I:I").ColumnWidth = 4.8
Columns("J:J").ColumnWidth = 4.5
Columns("K:L").ColumnWidth = 0.2
'Columns("L:L").ColumnWidth = 0.5
Columns("M:N").EntireColumn.AutoFit
Columns("O:O").ColumnWidth = 8.2
Columns("P:P").ColumnWidth = 4.1
Columns("Q:Q").ColumnWidth = 5.5
Columns("R:R").ColumnWidth = 2.3
Columns("S:T").ColumnWidth = 0.5
Columns("U:U").ColumnWidth = 3.8
Columns("V:W").ColumnWidth = 4.5
Columns("X:X").ColumnWidth = 12
Columns("Y:Y").ColumnWidth = 0.5
Columns("Z:Z").ColumnWidth = 1
Columns("AJ:AK").ColumnWidth = 4.5
Columns("AA:AI").EntireColumn.AutoFit
Rows(5).AutoFilter

'Tao Cacche1
Call Cacche1
ww = Sheets("Cacche1").Cells(5, 17)

Sheets("DuLieu").Select
'Xoa thu thuat khong phai YHCT-PHCN
    For i = 9 To 100
    str = Left(Cells(i, 2), 3)
        If str = "Y H" Then
             j = i
        End If
    Next
    For i = 1 To j - 7
    Rows(7).Delete
    Next
    str = ""
    j = 0
'Het xoa thu thuat khong phai YHCT-PHCN

Cells(1, 2) = ""
i = 7
While Cells(i, 5) <> ""
    'Xoa doi tuong <> BHYT
    str = Cells(i, 5)
        If str <> "" Then
        'If str = "BHYT" Then
        'Chuyen ten nguoi lam ve dung cot
            Cells(i, 35) = Left(Cells(i, 21), 50) & Left(Cells(i, 22), 50) & Left(Cells(i, 23), 50)
        'Doi ten thu thuat sang 21 ky tu
            Cells(i, 11) = Left(Cells(i, 8), 21)
        'Format thoi gian lam
            Cells(i, 15).FormulaR1C1 = "=DATE(YEAR(RC[-2]),MONTH(RC[-2]),DAY(RC[-2]))"
        'Time vào - time ra
            Cells(i, 17) = Cells(i, 14) - Cells(i, 13)
            Cells(i, 12).FormulaR1C1 = "=bo_dau_tieng_viet(RC[-1])"
            Cells(i, 24).FormulaR1C1 = "=bo_dau_tieng_viet(RC[11])"
        'GiaDv
            Cells(i, 9).FormulaR1C1 = "=VLOOKUP(RC[3], Cacche1!R1C2:R" & ww - 2 & "C8,7,0)"
        'Loai Dv Loai Nv
            Cells(i, 22).FormulaR1C1 = "=VLOOKUP(RC[-10], Cacche1!R1C2:R" & ww - 2 & "C3,2,0)"
            Cells(i, 23).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[1], Cacche1!R1C5:R6C6,2,0),""Ca2"")"
        'Add TimeBH
            Cells(i, 10).FormulaR1C1 = "=VLOOKUP(RC[2], Cacche1!R1C2:R" & ww - 2 & "C7,6,0)"
        'Loc T7/CN
            Cells(i, 18).FormulaR1C1 = "=IF(WEEKDAY(RC[-5])=1,""CN"",IF(WEEKDAY(RC[-5])=7,""T7"",IF(WEEKDAY(RC[-5])=6,""T6"", IF(WEEKDAY(RC[-5])=5,""T5"", IF(WEEKDAY(RC[-5])=4,""T4"", IF(WEEKDAY(RC[-5])=3,""T3"", IF(WEEKDAY(RC[-5])=2,""T2"")))))))"
        'Doi ten thu thuat cho pivot de nhin
            'Cells(i, 9).FormulaR1C1 = "=VLOOKUP(bo_dau_tieng_viet(Left(Cells(RC[-1], 21)), Cacche1!R1C2:R15C4,3,0)"
        'Chuyen Gio Phut bat dau lam
            Cells(i, 16).FormulaR1C1 = "=TIME(HOUR(RC[-2]),MINUTE(RC[-2]),SECOND(RC[-2]))"
        '00Sec
            Cells(i, 19).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
        '00Sec2
            Cells(i, 20).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
            
            i = i + 1
            m = i
        Else
                Rows(i).Delete
        End If
    'Het xoa doi tuong <> BHYT
Wend
i = 0
str = ""

Call PasteValueH(strx, m)
    
Range("AI7:AI" & m + 1).Value = "" 'Xoa Data
Range("U7:U" & m + 1).Value = "" 'Xoa Data
Range("K7:K" & m + 1).Value = "" 'Xoa Data

'Xoa nguoi lap bieu
    For i = 1 To 8
        Rows(m + 1).Delete
    Next
'Het Xoa nguoi lap bieu

'
Call CCHN_Error(m)

'Doi ten thu thuat cho pivot de nhin
Call ShortNameTT(strx, m, ww)

'Loc trung gio2
Application.Calculation = xlCalculationAutomatic 'Cái này de tinh Automatic moi coppy/paste duoc
       Cells(7, 27).Value = _
            "=COUNTIFS($B$7:$B$" & m - 1 & ",B7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7,$X$7:$X$" & m - 1 & ",""<>""&X7)+COUNTIFS($B$7:$B$" & m - 1 & ",B7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7)" 'Loc cot Phu second = 00
       Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
       Cells(7, 28).Value = _
            "=COUNTIFS($X$7:$X$" & m - 1 & ",X7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7,$B$7:$B$" & m - 1 & ",""<>""&B7)+COUNTIFS($X$7:$X$" & m - 1 & ",X7,$T$7:$T$" & m - 1 & ","">""&S7,$S$7:$S$" & m - 1 & ",""<""&T7)" 'Loc cot Phu second = 00
       Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
       'Sum TimeBH
       Cells(7, 21).Value = "=SUMIFS($J$7:$J$" & m - 1 & ",$X$7:$X$" & m - 1 & ",X7,$O$7:$O$" & m - 1 & ",O7)"
       Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
Application.Calculation = xlCalculationManual 'Het lenh Cái này de tinh Automatic moi coppy/paste duoc
    
    Range("AA7:AB" & m - 1).Copy 'Ctrl+Shift+D thì bo lenh copy 2lenh
    Range("AA7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Range("U7:U" & m - 1).Copy
    Range("U7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
'Hêt Loc trung gio2

'Tich nham nguoi khong truc
Call NoTruc(strx, m, ww)

'Color
Call CoLor(m)

'So Luong BN theo ngay
Call BN_Date(strx, m, ww)

'PivotH TT
    Sheets.Add.Name = "PivotH"
        
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "DuLieu!R5C1:R150000C35", Version:=6).CreatePivotTable TableDestination:= _
        "PivotH!R4C1", TableName:="PivotTableH", DefaultVersion:=6 'Mac dinh de range voi 150k rows
    Sheets("PivotH").Cells(4, 1).Select
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("TenNv")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("TenDv")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("NgayStart")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTableH").AddDataField ActiveSheet.PivotTables( _
        "PivotTableH").PivotFields("TenDv"), "Count of TenDv", xlCount
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("TenDv")
        .Orientation = xlColumnField 'Lân 2
        .Position = 1
    End With
'Hêt PivotH TT
    
'PivotH TimeBH
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "DuLieu!R5C1:R150000C35", Version:=6).CreatePivotTable TableDestination:= _
        "PivotH!R32C1", TableName:="PivotTableH2", DefaultVersion:=6 'Mac dinh de range voi 150k rows
    Sheets("PivotH").Cells(32, 1).Select
    With ActiveSheet.PivotTables("PivotTableH").PivotFields("NgayStart")
        .Orientation = xlPageField
        .Position = 1
    End With
        With ActiveSheet.PivotTables("PivotTableH2").PivotFields("NgayStart")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTableH2").PivotFields("TenNv")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTableH2").AddDataField ActiveSheet.PivotTables( _
        "PivotTableH2").PivotFields("TimeBH"), "Count of TimeBH", xlCount
    With ActiveSheet.PivotTables("PivotTableH2").PivotFields("Count of TimeBH")
        .Caption = "Sum of TimeBH"
        .Function = xlSum
    End With
    Range("B33:B55").NumberFormat = "[h]:mm;@"
'Hêt PivotH TimeBH
Cells(1, 5) = "Chu y ô: Blank"
    
Sheets("DuLieu").Select
Cells(1, 1).Select

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub ALL_FixTT() 'Ctrl+Shift+C
Attribute ALL_FixTT.VB_ProcData.VB_Invoke_Func = "C\n14"
Dim i As Variant
Dim str As String
Dim strx As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "So Phau thuat" Then
            strx = str
           Call FixTT_SoPhauthuatNew(strx)
        End If
        If str = "DuLieu" Then
            strx = str
           Call FixTT_DuLieuNew(strx)
        End If
Next
str = ""

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub FixTT_SoPhauthuatNew(strx As String) 'Co the chuyen bien so strx, m vao sub
Dim i As Variant
Dim m As Variant
Dim j As Variant
Dim a As Variant
Dim b As Variant
Dim x As Variant
Dim ww As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

ww = Sheets("Cacche1").Cells(5, 17)
       
Sheets("So Phau thuat").Select

i = 7
While Cells(i, 13) <> ""
            
        'Chuyen Gio Phut bat dau lam
            'Cells(i, 16).FormulaR1C1 = "=TIME(HOUR(RC[-2]),MINUTE(RC[-2]),SECOND(RC[-2]))"
        '00Sec
            Cells(i, 19).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
        '00Sec2
            Cells(i, 20).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
                    
            i = i + 1
            m = i
Wend
i = 0
'Range("A7", Range("AK" & m - 1).End(xlUp)).Sort [O7], xlAscending, Header:=xlYes 'Loc Date
Range("A7:AK" & m - 1).Sort Key1:=Range("O7"), Order1:=xlAscending, Header:=xlNo

'Loc trung gio2
j = Sheets("Cacche1").Cells(ww + 1, 4)
a = 0
b = 0
x = 0
For x = 1 To j
    b = a
    a = a + Sheets("Cacche1").Cells(ww + x, 7)
    If a + 6 <= m - 1 Then
        If x = 1 Then
            For i = 7 To a + 6
                   Cells(i, 27).Value = _
                        "=COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$Y$" & b + 7 & ":$Y$" & a + 6 & ",""<>""&Y" & i & ")+COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
                   Cells(i, 28).Value = _
                        "=COUNTIFS($Y$" & b + 7 & ":$Y$" & a + 6 & ",Y" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$B$" & b + 7 & ":$B$" & a + 6 & ",""<>""&B" & i & ")+COUNTIFS($Y$" & b + 7 & ":$Y$" & a + 6 & ",Y" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
                   'Sum Time BH
                   Cells(i, 21).Value = "=SUMIFS($K$" & i & ":$K$" & a + 6 & ",$Y$" & i & ":$Y$" & a + 6 & ",Y" & i & ",$O$" & i & ":$O$" & a + 6 & ",O" & i & ")"
                   'Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
                   'Cacche
                   Cells(i, 35).Value = "=AA" & i & "+AB" & i & ""
                   'Range("AI7").AutoFill Destination:=Range("AI7:AI" & m - 1)
            Next
        Else
            For i = (b + 7) To a + 6
                   Cells(i, 27).Value = _
                        "=COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$Y$" & b + 7 & ":$Y$" & a + 6 & ",""<>""&Y" & i & ")+COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
                   Cells(i, 28).Value = _
                        "=COUNTIFS($Y$" & b + 7 & ":$Y$" & a + 6 & ",Y" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$B$" & b + 7 & ":$B$" & a + 6 & ",""<>""&B" & i & ")+COUNTIFS($Y$" & b + 7 & ":$Y$" & a + 6 & ",Y" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
                   'Sum Time BH
                   Cells(i, 21).Value = "=SUMIFS($K$" & i & ":$K$" & a + 6 & ",$Y$" & i & ":$Y$" & a + 6 & ",Y" & i & ",$O$" & i & ":$O$" & a + 6 & ",O" & i & ")"
                   'Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
                   'Cacche
                   Cells(i, 35).Value = "=AA" & i & "+AB" & i & ""
                   'Range("AI7").AutoFill Destination:=Range("AI7:AI" & m - 1)
            Next
        
        End If
        'Hêt Loc trung gio2
    End If
Next
a = ""
b = ""
x = ""
j = ""

Application.Calculation = xlCalculationAutomatic

'FixTT
Cells(5, 35) = "00:05" 'Step
Cells(5, 36) = "07:30" 'Gio mua dong mua he
Cells(5, 37) = "13:00"

For i = 7 To m
    If Cells(i, 35) > 2 Then
        Cells(i, 32) = "Fix"
        Cells(i, 14).FormulaR1C1 = "=RC[-1] + RC[-3]"
        For j = 0 To 120   'Làm den 19h00
            If Cells(i, 35) > 2 Then
                If j = 0 Then
                        Cells(i, 13) = Cells(i, 15) + Cells(5, 36)
                Else
                        If j = 55 Then
                            Cells(i, 13) = Cells(i, 15) + Cells(5, 37)
                        Else
                            Cells(i, 13) = Cells(i, 13) + Cells(5, 35)
                        End If
                End If
            End If
        Next
    End If
Next
'Hêt FixTT

'Sum TimeBH
Cells(7, 21).Value = "=SUMIFS($K$7:$K$" & m - 1 & ",$Y$7:$Y$" & m - 1 & ",Y7,$O$7:$O$" & m - 1 & ",O7)"
Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
       
Call PasteValueH(strx, m)

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub FixTT_DuLieuNew(strx As String) 'Co the chuyen bien so strx, m vao sub
Dim i As Variant
Dim m As Variant
Dim j As Variant
Dim a As Variant
Dim b As Variant
Dim x As Variant
Dim ww As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

ww = Sheets("Cacche1").Cells(5, 17)

Sheets("DuLieu").Select

i = 7
While Cells(i, 13) <> ""
            
        'Chuyen Gio Phut bat dau lam
            'Cells(i, 16).FormulaR1C1 = "=TIME(HOUR(RC[-2]),MINUTE(RC[-2]),SECOND(RC[-2]))"
        '00Sec
            Cells(i, 19).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
        '00Sec2
            Cells(i, 20).FormulaR1C1 = _
                    "=DATE(YEAR(RC[-6]),MONTH(RC[-6]),DAY(RC[-6]))+TIME(HOUR(RC[-6]),MINUTE(RC[-6]),SECOND(R1C2))"
                    
            i = i + 1
            m = i
Wend
i = 0
'Range("A7", Range("AB" & m - 1).End(xlUp)).Sort [O7], xlAscending, Header:=xlNo 'Loc Date
Range("A7:AB" & m - 1).Sort Key1:=Range("O7"), Order1:=xlAscending, Header:=xlNo

'Loc trung gio2
j = Sheets("Cacche1").Cells(ww + 1, 4)
a = 0
b = 0
x = 0
For x = 1 To j
    b = a
    a = a + Sheets("Cacche1").Cells(ww + 1 + x, 7)
    If a + 6 <= m - 1 Then
        If x = 1 Then
            For i = 7 To a + 6
                   Cells(i, 27).Value = _
                        "=COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$X$" & b + 7 & ":$X$" & a + 6 & ",""<>""&X" & i & ")+COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
                   Cells(i, 28).Value = _
                        "=COUNTIFS($X$" & b + 7 & ":$X$" & a + 6 & ",X" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$B$" & b + 7 & ":$B$" & a + 6 & ",""<>""&B" & i & ")+COUNTIFS($X$" & b + 7 & ":$X$" & a + 6 & ",X" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
                   'Sum Time BH
                   Cells(i, 21).Value = "=SUMIFS($J$" & i & ":$J$" & a + 6 & ",$X$" & i & ":$X$" & a + 6 & ",X" & i & ",$O$" & i & ":$O$" & a + 6 & ",O" & i & ")"
                   'Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
                   'Cacche
                   Cells(i, 35).Value = "=AA" & i & "+AB" & i & ""
                   'Range("AI7").AutoFill Destination:=Range("AI7:AI" & m - 1)
            Next
        Else
            For i = (b + 7) To a + 6
                   Cells(i, 27).Value = _
                        "=COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$X$" & b + 7 & ":$X$" & a + 6 & ",""<>""&X" & i & ")+COUNTIFS($B$" & b + 7 & ":$B$" & a + 6 & ",B" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AA7").AutoFill Destination:=Range("AA7:AA" & m - 1)
                   Cells(i, 28).Value = _
                            "=COUNTIFS($X$" & b + 7 & ":$X$" & a + 6 & ",X" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ",$B$" & b + 7 & ":$B$" & a + 6 & ",""<>""&B" & i & ")+COUNTIFS($X$" & b + 7 & ":$X$" & a + 6 & ",X" & i & ",$T$" & b + 7 & ":$T$" & a + 6 & ","">""&S" & i & ",$S$" & b + 7 & ":$S$" & a + 6 & ",""<""&T" & i & ")" 'Loc cot Phu second = 00
                   'Range("AB7").AutoFill Destination:=Range("AB7:AB" & m - 1)
                   'Sum Time BH
                   Cells(i, 21).Value = "=SUMIFS($K$" & i & ":$K$" & a + 6 & ",$X$" & i & ":$X$" & a + 6 & ",X" & i & ",$O$" & i & ":$O$" & a + 6 & ",O" & i & ")"
                   'Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)
                   'Cacche
                   Cells(i, 35).Value = "=AA" & i & "+AB" & i & ""
                   'Range("AI7").AutoFill Destination:=Range("AI7:AI" & m - 1)
            Next
        
        End If
        'Hêt Loc trung gio2
    End If
Next
a = ""
b = ""
x = ""
j = ""

Application.Calculation = xlCalculationAutomatic

'FixTT
Cells(4, 35) = "00:05" 'Step
Cells(4, 36) = "07:30" 'Gio mua dong mua he
Cells(4, 37) = "13:00"

For i = 7 To m
    If Cells(i, 35) > 2 Then
        Cells(i, 32) = "Fix"
        Cells(i, 14).FormulaR1C1 = "=RC[-1] + RC[-4]"
        For j = 0 To 120   'Làm den 19h00
            If Cells(i, 35) > 2 Then
                If j = 0 Then
                        Cells(i, 13) = Cells(i, 15) + Cells(4, 36)
                Else
                        If j = 55 Then
                            Cells(i, 13) = Cells(i, 15) + Cells(4, 37)
                        Else
                            Cells(i, 13) = Cells(i, 13) + Cells(4, 35)
                        End If
                End If
            End If
        Next
    End If
Next
'Hêt FixTT

'Sum TimeBH
Cells(7, 21).Value = "=SUMIFS($J$7:$J$" & m - 1 & ",$X$7:$X$" & m - 1 & ",X7,$O$7:$O$" & m - 1 & ",O7)"
Range("U7").AutoFill Destination:=Range("U7:U" & m - 1)

Call PasteValueH(strx, m)

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub NoTruc(strx As String, m As Variant, ww As Integer) 'Tich nham nguoi khong truc
Dim str As String
Dim i As Integer

For i = 1 To Sheets.Count
str = Sheets(i).Name
    If str = "Truc" Then
        Sheets("Truc").Select
        While Cells(1, 1) = "" And Cells(1, 2) = "" And Cells(1, 3) = ""
            Rows(1).Delete
        Wend
            Range("A1:B500").Copy
        Sheets("Cacche1").Select
            Range("K" & ww + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Sheets(strx).Select
Application.Calculation = xlCalculationAutomatic 'Cái này de tinh Automatic moi coppy/paste duoc

            Cells(7, 31).Value = "=VLOOKUP(X7,Cacche1!$H$" & ww + 1 & ":$I$" & ww + 40 & ",2,FALSE)"
                Range("AE7").AutoFill Destination:=Range("AE7:AE" & m - 1)
            Cells(7, 30).Value = "=VLOOKUP(O7,Cacche1!$K$" & ww + 1 & ":$L$" & ww + 500 & ",2,FALSE)"
                Range("AD7").AutoFill Destination:=Range("AD7:AD" & m - 1)
            Cells(7, 29).Value = "=IFERROR(SEARCH(AE7,AD7,1),""No"")"
                Range("AC7").AutoFill Destination:=Range("AC7:AC" & m - 1)
            Cells(7, 35).Value = "=IF(AC7<>""No"",""Yes"",IF(AND(AC7=""No"",OR(R7=""T7"",R7=""CN"")),""NoT7CN"",AC7))" 'Cacche
                Range("AI7").AutoFill Destination:=Range("AI7:AI" & m - 1)
            
Application.Calculation = xlCalculationManual 'Het lenh Cái này de tinh Automatic moi coppy/paste duoc

        Range("AI7:AI" & m - 1).Copy
        Range("AC7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Range("AI7:AI" & m - 1).Value = "" 'Xoa cacche
            
        Range("AC7:AE" & m - 1).Copy
        Range("AC7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
     End If
 Next
str = ""
'Hét Tich nham nguoi khong truc
End Sub

Sub CoLor(m As Variant)
    
    With Range("AA6:AB" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    With Range("Q6:Q" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    With Range("X6:X" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    
    With Range("I6:I" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .CoLor = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With Range("M6:N" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    With Range("H6:H" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
    
    With Range("U6:U" & m - 1).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .CoLor = 15773696
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
'Hêt Color
End Sub

Sub Cacche1()
'Application.DisplayAlerts = False
'Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

'Loc CCHN YHCT-PHCN
Sheets.Add.Name = "Cacche1"
Sheets("Cacche1").Move Before:=Sheets(3)
Sheets("Cacche1").Select
Cells(1, 2) = "Cuu (ngai cuu, tui ch"
Cells(2, 2) = "Dien cham"
Cells(3, 2) = "Ngam thuoc YHCT bo ph"
Cells(4, 2) = "Sac thuoc thang"
Cells(5, 2) = "Thuy cham (Chua bao g"
Cells(6, 2) = "Xoa bop bam huyet ban"
Cells(7, 2) = "Dieu tri bang cac don"
Cells(8, 2) = "Dieu tri bang may keo"
Cells(9, 2) = "Dieu tri bang Parafin"
Cells(10, 2) = "Dieu tri bang sieu am"
Cells(11, 2) = "Tap van dong co tro g"
Cells(12, 2) = "Tap van dong thu dong"
Cells(13, 2) = "Dien cham khong kim 2"
Cells(13, 2) = "Dien cham khong kim 2"
Cells(14, 2) = "Tap len, xuong cau th"
Cells(15, 2) = "Dieu tri bang tia hon"
Cells(16, 2) = "Tap ngoi thang bang t"
Cells(17, 2) = "Ky thuat xoa bop vung"
Cells(1, 3) = "YHCT"
Cells(2, 3) = "YHCT"
Cells(3, 3) = "_"
Cells(4, 3) = "YHCT"
Cells(5, 3) = "YHCT"
Cells(6, 3) = "YHCT"
Cells(7, 3) = "PHCN"
Cells(8, 3) = "PHCN"
Cells(9, 3) = "PHCN"
Cells(10, 3) = "PHCN"
Cells(11, 3) = "PHCN"
Cells(12, 3) = "PHCN"
Cells(13, 3) = "_"
Cells(14, 3) = "PHCN"
Cells(15, 3) = "PHCN"
Cells(16, 3) = "PHCN"
Cells(17, 3) = "YHCT"

Cells(1, 4) = "CuuNgai"
Cells(2, 4) = "DienCham"
Cells(3, 4) = "NgamChan"
Cells(4, 4) = "SacThuoc"
Cells(5, 4) = "ThuyCham"
Cells(6, 4) = "XoaBop"
Cells(7, 4) = "DienXung"
Cells(8, 4) = "KeoGian"
Cells(9, 4) = "Parafin"
Cells(10, 4) = "SieuÂm"
Cells(11, 4) = "TapVD"
Cells(12, 4) = "TapVD"
Cells(13, 4) = "KoKim"
Cells(14, 4) = "TapPHCN"
Cells(15, 4) = "Hongoai"
Cells(16, 4) = "TapVD"
Cells(17, 4) = "XoaBop"

Cells(1, 7) = "00:15"  '"Cuu ngai"
Cells(2, 7) = "00:20"   '"Dien cham"
Cells(3, 7) = "00:0"   '"Ngam chan"
Cells(4, 7) = "00:5"   '"Sac thuoc"
Cells(5, 7) = "00:5"   '"Thuy cham"
Cells(6, 7) = "00:15"  '"Xoa bop"
Cells(7, 7) = "00:15"  '"Dien xung"
Cells(8, 7) = "00:5"   '"Keo gian"
Cells(9, 7) = "00:15"   '"Parafin"
Cells(10, 7) = "00:15" '"Sieu am"
Cells(11, 7) = "00:20" '"Tap VD"
Cells(12, 7) = "00:20" '"Tap VD"
Cells(13, 7) = "00:0"  '"Ko Kim"
Cells(14, 7) = "00:10"  '"Tap PHCN"
Cells(15, 7) = "00:5"  'Hongoai
Cells(16, 7) = "00:20" 'Tap VD
Cells(17, 7) = "00:15"  '"Xoa bop"

Cells(1, 8) = "35500" ' Cuu ngai"
Cells(2, 8) = "67300" '"Dien cham"
Cells(3, 8) = "1" '"Ngam chan"
Cells(4, 8) = "12500" '"Sac thuoc"
Cells(5, 8) = "66100" '"Thuy cham"
Cells(6, 8) = "65500" '"Xoa bop"
Cells(7, 8) = "41400" '"Dien xung"
Cells(8, 8) = "45800" '"Keo gian"
Cells(9, 8) = "42400" '"Parafin"
Cells(10, 8) = "45600" '"Sieu am"
Cells(11, 8) = "46900" '"Tap VD"
Cells(12, 8) = "45600" '"Tap VD"
Cells(13, 8) = "1" '"Ko Kim"
Cells(14, 8) = "29000" '"TapPHCN"
Cells(15, 8) = "35200" '"Hongoai"
Cells(16, 8) = "46900" '"Tap VD"
Cells(17, 8) = "41800" '"Xoa bop"

'Cells(1, 5) = "Tran Thi Hong Thinh"
'Cells(2, 5) = "Tran Thi Lan"
'Cells(3, 5) = "Nguyen Thi Thu Thao"
'Cells(4, 5) = "Le Thi Hong Lien"
'Cells(5, 5) = "Luong Thi Minh Nguyet"
Cells(6, 5) = "Hoang Thi Nhu Quynh"

Cells(1, 6) = "PHCN" '"Tran Thi Hong Thinh"
Cells(2, 6) = "PHCN" '"Tran Thi Lan"
Cells(3, 6) = "PHCN" '"Nguyen Thi Thu Thao"
Cells(4, 6) = "YHCT" '"Le Thi Hong Lien"
Cells(5, 6) = "YHCT" '"Luong Thi Minh Nguyet"
Cells(6, 6) = "YHCT" '"Hoang Thi Nhu Quynh"

'Nhac nho
Cells(2, 13) = "Pivot: Thua Nguoi so voi bang cham cong, Tich nham nguoi khoa khác, SumThuThuat/People, Blank"
Cells(3, 13) = "DuLieu: YHCT-PHCN, Time Vao-Ra, T7/CN, Trung Gio,"
Cells(4, 13) = "BS khám 8min/1BN (Thuong la cuoi tuan co it ngTruc)"
Cells(5, 13) = "Tham so dich chuyen BN/Date (ww):"
Cells(5, 17) = 20

Call Cacche11

'Application.DisplayAlerts = True
'Application.ScreenUpdating = True
'Application.Calculation = xlCalculationAutomatic
End Sub
Sub Cacche11()
Dim ww As Integer

Sheets("Cacche1").Select

ww = Sheets("Cacche1").Cells(5, 17)

Range("K" & ww + 1 & ":K500").NumberFormat = "dd/mm/yyyy;@"
Columns("K:K").ColumnWidth = 10

Cells(ww, 2) = "Date"
Cells(ww, 3) = "BN/Date"
Cells(ww, 4) = "SumDate"
Cells(ww, 5) = "EndDate"
Cells(ww, 6) = "T7/CN"
Cells(ww, 7) = "TT/Date"
Cells(ww, 8) = "TenNV"
Cells(ww, 9) = "TenTruc"
Cells(ww, 11) = "Ngay Truc"
Cells(ww, 12) = "Nguoi Truc"
Cells(ww, 14) = "Dr/Date"
Cells(ww, 15) = "TimeKham"

Cells(ww + 1, 8) = "BS CKI. Pham Van Anh"
Cells(ww + 2, 8) = "BS. Phung Manh Dat"
Cells(ww + 3, 8) = "BS. Nguyen Thi Huyen"
Cells(ww + 4, 8) = "BS. Vu Thi Ngoc"
Cells(ww + 5, 8) = "BS. Nguyen Xuan Hoang"
Cells(ww + 6, 8) = "BS. Nguyen Thuy Linh"
Cells(ww + 7, 8) = "BS. Nguyen Thi Le Thu"
Cells(ww + 8, 8) = "BS. Le Thi Nhu Hoan"
Cells(ww + 9, 8) = "Hoang Thi Hien"
Cells(ww + 10, 8) = "Dao Ngoc Cuong"
Cells(ww + 11, 8) = "Nguyen Van Dinh"
Cells(ww + 12, 8) = "Pham Thi Ha"
Cells(ww + 13, 8) = "Tran Thi Lan"
Cells(ww + 14, 8) = "Le Thi Hong Lien"
Cells(ww + 15, 8) = "Phung Thi Nhung"
Cells(ww + 16, 8) = "Nguyen Thi Thu Thao"
Cells(ww + 17, 8) = "Tran Thi Kim Thuy"
Cells(ww + 18, 8) = "Tran Thi Hong Thinh"
Cells(ww + 19, 8) = "Luong Thi Minh Nguyet"

Cells(ww + 1, 9) = "Pham Van Anh"
Cells(ww + 2, 9) = "Phung Manh Dat"
Cells(ww + 3, 9) = "Nguyen Thi Huyen"
Cells(ww + 4, 9) = "Vu Thi Ngoc"
Cells(ww + 5, 9) = "Nguyen Xuan Hoang"
Cells(ww + 6, 9) = "Nguyen Thuy Linh"
Cells(ww + 7, 9) = "Nguyen Thi Le Thu"
Cells(ww + 8, 9) = "Le Thi Nhu Hoan"
Cells(ww + 9, 9) = "Hoang Thi Hien"
Cells(ww + 10, 9) = "Dao Ngoc Cuong"
Cells(ww + 11, 9) = "Nguyen Van Dinh"
Cells(ww + 12, 9) = "Pham Thi Ha"
Cells(ww + 13, 9) = "Tran Thi Lan"
Cells(ww + 14, 9) = "Le Thi Hong Lien"
Cells(ww + 15, 9) = "Phung Thi Nhung"
Cells(ww + 16, 9) = "Nguyen Thi Thu Thao"
Cells(ww + 17, 9) = "Tran Thi Kim Thuy"
Cells(ww + 18, 9) = "Tran Thi Hong Thinh"
Cells(ww + 19, 9) = "Luong Thi Minh Nguyet"

End Sub

Sub BN_Date(strx As String, m As Variant, ww As Integer) 'So Luong BN theo ngay
Dim i As Integer

Sheets(strx).Select
    Range("B7:B" & m + 1).Copy
Sheets("Cacche1").Select
    Range("Y" & ww + 1).Select
    ActiveSheet.Paste
Sheets(strx).Select
    Range("O7:O" & m + 1).Copy
Sheets("Cacche1").Select
    Range("Z" & ww + 1).Select
    ActiveSheet.Paste

Range("Y" & ww + 1 & ":Z" & (m + 25)).Select
    Application.CutCopyMode = False
    ActiveSheet.Range("$Y$" & ww + 1 & ":$Z$" & (m + ww + 10)).RemoveDuplicates Columns:=Array(1, 2), Header:=xlYes
        
'Application.Calculation = xlCalculationAutomatic
Cells(ww + 1, 2).Formula = "=MIN(Z" & ww + 1 & ":Z150000)"
Cells(ww + 1, 5).Formula = "=MAX(Z" & ww + 1 & ":Z150000)"
Cells(ww + 1, 4) = Cells(ww + 1, 5) - Cells(ww + 1, 2) + 1
    If Cells(ww + 1, 4) > 1 Then 'Them 1 BN thieu khi load nhieu hon 1 ngay
                For i = ww + 2 To Cells(ww + 1, 4) + ww
                    Cells(i, 2) = Cells(i - 1, 2) + 1
                    Cells(i, 3).Formula = "=COUNTIF($Z$" & ww + 1 & ":Z150000,B" & i & ")+1"
                    'Loc T7/CN sheet Cacche1
                    'Cells(i, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",""_""))"
                    Cells(i, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",IF(WEEKDAY(RC[-4])=6,""T6"", IF(WEEKDAY(RC[-4])=5,""T5"", IF(WEEKDAY(RC[-4])=4,""T4"", IF(WEEKDAY(RC[-4])=3,""T3"", IF(WEEKDAY(RC[-4])=2,""T2"")))))))"
                Next
        Cells(ww + 1, 3).Formula = "=COUNTIF($Z$" & ww + 1 & ":Z150000,B" & ww + 1 & ")+1"
        Cells(ww + 1, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",IF(WEEKDAY(RC[-4])=6,""T6"", IF(WEEKDAY(RC[-4])=5,""T5"", IF(WEEKDAY(RC[-4])=4,""T4"", IF(WEEKDAY(RC[-4])=3,""T3"", IF(WEEKDAY(RC[-4])=2,""T2"")))))))"
        
    Else 'Khong them 1 BN thieu khi load 1 ngay
                'For i = ww + 2 To Cells(ww + 1, 4) + ww
                '    Cells(i, 2) = Cells(i - 1, 2) + 1
                '    Cells(i, 3).Formula = "=COUNTIF($Z$" & ww + 1 & ":Z150000,B" & i & ")"
                    'Loc T7/CN sheet Cacche1
                    'Cells(i, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",""_""))"
                '    Cells(i, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",IF(WEEKDAY(RC[-4])=6,""T6"", IF(WEEKDAY(RC[-4])=5,""T5"", IF(WEEKDAY(RC[-4])=4,""T4"", IF(WEEKDAY(RC[-4])=3,""T3"", IF(WEEKDAY(RC[-4])=2,""T2"")))))))"
                'Next
        Cells(ww + 1, 3).Formula = "=COUNTIF($Z$" & ww + 1 & ":Z150000,B" & ww + 1 & ")"
        Cells(ww + 1, 6).FormulaR1C1 = "=IF(WEEKDAY(RC[-4])=1,""CN"",IF(WEEKDAY(RC[-4])=7,""T7"",IF(WEEKDAY(RC[-4])=6,""T6"", IF(WEEKDAY(RC[-4])=5,""T5"", IF(WEEKDAY(RC[-4])=4,""T4"", IF(WEEKDAY(RC[-4])=3,""T3"", IF(WEEKDAY(RC[-4])=2,""T2"")))))))"
    End If
'SumTT/Date
i = ww + 1
While Cells(i, 2) <> ""
    Cells(i, 7).Formula = "=COUNTIFS('" & strx & "'!$O$7:$O$" & m & ",Cacche1!B" & i & ")"
    i = i + 1
Wend
i = 0

Range("B" & ww + 1 & ":G" & Cells(ww + 1, 4) + ww + 2).Copy
Range("B" & ww + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

Columns(25).Delete 'Xoa cot Phu
Columns(25).Delete
Columns("B:E").ColumnWidth = 12
Cells(1, 1).Select
'Het So Luong BN theo ngay
End Sub

Sub ShortNameTT(strx As String, m As Variant, ww As Integer) 'Doi ten thu thuat cho pivot de nhin
Dim i As Integer

If strx = "So Phau thuat" Then
        For i = 7 To m - 1
           Cells(i, 35).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-26], Cacche1!R1C2:R" & ww - 2 & "C4,3,0),""NoList"")"
        Next
End If
If strx = "DuLieu" Then
        For i = 7 To m - 1
           Cells(i, 35).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-23], Cacche1!R1C2:R" & ww - 2 & "C4,3,0),""NoList"")"
        Next
End If

'Thu thuat khong co ten trong list cacche1
Application.Calculation = xlCalculationAutomatic
If strx = "So Phau thuat" Then
    For i = 7 To m - 1
        If Cells(i, 35) = "NoList" Then
              Cells(i, 35).FormulaR1C1 = "=bo_dau_tieng_viet(RC[-23])"
        End If
    Next
    Range("AI7:AI" & m).Copy 'Cacche colunm
    Range("I7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End If
If strx = "DuLieu" Then
    For i = 7 To m - 1
        If Cells(i, 35) = "NoList" Then 'Thu thuat khong co ten trong list cacche1
              Cells(i, 35).FormulaR1C1 = "=bo_dau_tieng_viet(RC[-23])"
        End If
    Next
    Range("AI7:AI" & m).Copy 'Cacche colunm
    Range("H7").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
End If
Application.Calculation = xlCalculationManual
'Hét Thu thuat khong co ten trong list cacche1
  
    Range("AI7:AI" & m + 1).Value = "" 'Xoa Data
    Range("L7:L" & m + 1).Value = ""
'Het Doi ten thu thuat cho pivot de nhin
End Sub

Sub XuLyBangTruc()  'Ctrl+Shift+Y
Attribute XuLyBangTruc.VB_ProcData.VB_Invoke_Func = "Y\n14"
Dim i As Integer
Dim j As Variant
Dim x As Variant

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

For x = 1 To 12
Sheets(x).Select
    For j = 4 To 35
        For i = 9 To 50
            If Cells(i, j) = 1 Then
                Cells(i, j) = Cells(i, 2)
            End If
        Next
    Next
Next

Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub

Sub PasteValueH(strx As String, m As Variant)
Sheets(strx).Range("A6:AK" & m).Copy
Sheets(strx).Range("A6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Sheets(strx).Cells(1, 1).Select
End Sub

Sub CCHN_Error(m As Variant)
Dim i As Integer
For i = 7 To m
    On Error Resume Next
     If Cells(i, 22) = "YHCT" Then
        If Cells(i, 23) = "PHCN" Then
            Cells(i, 33) = "Sai"
        End If
     End If
     If Cells(i, 22) = "PHCN" Then
        If Cells(i, 23) = "YHCT" Then
            Cells(i, 33) = "Sai"
        End If
     End If
Next
End Sub

Function bo_dau_tieng_viet(Text As String) As String
  Dim AsciiDict As Object
  Set AsciiDict = CreateObject("scripting.dictionary")
  AsciiDict(192) = "A"
  AsciiDict(193) = "A"
  AsciiDict(194) = "A"
  AsciiDict(195) = "A"
  AsciiDict(196) = "A"
  AsciiDict(197) = "A"
  AsciiDict(199) = "C"
  AsciiDict(200) = "E"
  AsciiDict(201) = "E"
  AsciiDict(202) = "E"
  AsciiDict(203) = "E"
  AsciiDict(204) = "I"
  AsciiDict(205) = "I"
  AsciiDict(206) = "I"
  AsciiDict(207) = "I"
  AsciiDict(208) = "D"
  AsciiDict(209) = "N"
  AsciiDict(210) = "O"
  AsciiDict(211) = "O"
  AsciiDict(212) = "O"
  AsciiDict(213) = "O"
  AsciiDict(214) = "O"
  AsciiDict(217) = "U"
  AsciiDict(218) = "U"
  AsciiDict(219) = "U"
  AsciiDict(220) = "U"
  AsciiDict(221) = "Y"
  AsciiDict(224) = "a"
  AsciiDict(225) = "a"
  AsciiDict(226) = "a"
  AsciiDict(227) = "a"
  AsciiDict(228) = "a"
  AsciiDict(229) = "a"
  AsciiDict(231) = "c"
  AsciiDict(232) = "e"
  AsciiDict(233) = "e"
  AsciiDict(234) = "e"
  AsciiDict(235) = "e"
  AsciiDict(236) = "i"
  AsciiDict(237) = "i"
  AsciiDict(238) = "i"
  AsciiDict(239) = "i"
  AsciiDict(240) = "d"
  AsciiDict(241) = "n"
  AsciiDict(242) = "o"
  AsciiDict(243) = "o"
  AsciiDict(244) = "o"
  AsciiDict(245) = "o"
  AsciiDict(246) = "o"
  AsciiDict(249) = "u"
  AsciiDict(250) = "u"
  AsciiDict(251) = "u"
  AsciiDict(252) = "u"
  AsciiDict(253) = "y"
  AsciiDict(255) = "y"
  AsciiDict(352) = "S"
  AsciiDict(353) = "s"
  AsciiDict(376) = "Y"
  AsciiDict(381) = "Z"
  AsciiDict(382) = "z"
  AsciiDict(258) = "A"
  AsciiDict(259) = "a"
  AsciiDict(272) = "D"
  AsciiDict(273) = "d"
  AsciiDict(296) = "I"
  AsciiDict(297) = "i"
  AsciiDict(360) = "U"
  AsciiDict(361) = "u"
  AsciiDict(416) = "O"
  AsciiDict(417) = "o"
  AsciiDict(431) = "U"
  AsciiDict(432) = "u"
  AsciiDict(7840) = "A"
  AsciiDict(7841) = "a"
  AsciiDict(7842) = "A"
  AsciiDict(7843) = "a"
  AsciiDict(7844) = "A"
  AsciiDict(7845) = "a"
  AsciiDict(7846) = "A"
  AsciiDict(7847) = "a"
  AsciiDict(7848) = "A"
  AsciiDict(7849) = "a"
  AsciiDict(7850) = "A"
  AsciiDict(7851) = "a"
  AsciiDict(7852) = "A"
  AsciiDict(7853) = "a"
  AsciiDict(7854) = "A"
  AsciiDict(7855) = "a"
  AsciiDict(7856) = "A"
  AsciiDict(7857) = "a"
  AsciiDict(7858) = "A"
  AsciiDict(7859) = "a"
  AsciiDict(7860) = "A"
  AsciiDict(7861) = "a"
  AsciiDict(7862) = "A"
  AsciiDict(7863) = "a"
  AsciiDict(7864) = "E"
  AsciiDict(7865) = "e"
  AsciiDict(7866) = "E"
  AsciiDict(7867) = "e"
  AsciiDict(7868) = "E"
  AsciiDict(7869) = "e"
  AsciiDict(7870) = "E"
  AsciiDict(7871) = "e"
  AsciiDict(7872) = "E"
  AsciiDict(7873) = "e"
  AsciiDict(7874) = "E"
  AsciiDict(7875) = "e"
  AsciiDict(7876) = "E"
  AsciiDict(7877) = "e"
  AsciiDict(7878) = "E"
  AsciiDict(7879) = "e"
  AsciiDict(7880) = "I"
  AsciiDict(7881) = "i"
  AsciiDict(7882) = "I"
  AsciiDict(7883) = "i"
  AsciiDict(7884) = "O"
  AsciiDict(7885) = "o"
  AsciiDict(7886) = "O"
  AsciiDict(7887) = "o"
  AsciiDict(7888) = "O"
  AsciiDict(7889) = "o"
  AsciiDict(7890) = "O"
  AsciiDict(7891) = "o"
  AsciiDict(7892) = "O"
  AsciiDict(7893) = "o"
  AsciiDict(7894) = "O"
  AsciiDict(7895) = "o"
  AsciiDict(7896) = "O"
  AsciiDict(7897) = "o"
  AsciiDict(7898) = "O"
  AsciiDict(7899) = "o"
  AsciiDict(7900) = "O"
  AsciiDict(7901) = "o"
  AsciiDict(7902) = "O"
  AsciiDict(7903) = "o"
  AsciiDict(7904) = "O"
  AsciiDict(7905) = "o"
  AsciiDict(7906) = "O"
  AsciiDict(7907) = "o"
  AsciiDict(7908) = "U"
  AsciiDict(7909) = "u"
  AsciiDict(7910) = "U"
  AsciiDict(7911) = "u"
  AsciiDict(7912) = "U"
  AsciiDict(7913) = "u"
  AsciiDict(7914) = "U"
  AsciiDict(7915) = "u"
  AsciiDict(7916) = "U"
  AsciiDict(7917) = "u"
  AsciiDict(7918) = "U"
  AsciiDict(7919) = "u"
  AsciiDict(7920) = "U"
  AsciiDict(7921) = "u"
  AsciiDict(7922) = "Y"
  AsciiDict(7923) = "y"
  AsciiDict(7924) = "Y"
  AsciiDict(7925) = "y"
  AsciiDict(7926) = "Y"
  AsciiDict(7927) = "y"
  AsciiDict(7928) = "Y"
  AsciiDict(7929) = "y"
  AsciiDict(8363) = "d"
  Text = Trim(Text)
  If Text = "" Then Exit Function
  Dim Char As String, _
    NormalizedText As String, _
    UnicodeCharCode As Long, _
    i As Long
  'Remove accent marks (diacritics) from text
  For i = 1 To Len(Text)
    Char = Mid(Text, i, 1)
    UnicodeCharCode = AscW(Char)
    If (UnicodeCharCode < 0) Then
      'See http://support.microsoft.com/kb/272138
      UnicodeCharCode = 65536 + UnicodeCharCode
    End If
    If AsciiDict.Exists(UnicodeCharCode) Then
      NormalizedText = NormalizedText & AsciiDict.Item(UnicodeCharCode)
    Else
      NormalizedText = NormalizedText & Char
    End If
  Next
  bo_dau_tieng_viet = NormalizedText
End Function

Sub GioMuaDongHe()
'SoPhauThuat
Cells(5, 35) = "00:05" 'Step
Cells(5, 36) = "07:30" 'Gio mua dong mua he
Cells(5, 37) = "13:00"

'DuLieu
Cells(4, 35) = "00:05" 'Step
Cells(4, 36) = "07:30" 'Gio mua dong mua he
Cells(4, 37) = "13:00"

End Sub
Sub CopyTime()  'Ctrl+Shift+B
Attribute CopyTime.VB_ProcData.VB_Invoke_Func = "B\n14"
Dim i As Variant
Dim str As String
Dim m As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

m = Sheets("So Phau thuat (2)").UsedRange.Rows(Sheets("So Phau thuat (2)").UsedRange.Rows.Count).Row

For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "So Phau thuat" Then
           Sheets("So Phau thuat (2)").Cells(7, 10).Formula = "=VLOOKUP('So Phau thuat (2)'!A7,'So Phau thuat'!$A$7:$N$" & m & ",13,TRUE)"
           Range("J7").AutoFill Destination:=Range("J7:J" & m)
           Sheets("So Phau thuat (2)").Cells(7, 11).Formula = "=VLOOKUP('So Phau thuat (2)'!A7,'So Phau thuat'!$A$7:$N$" & m & ",14,TRUE)"
           Range("K7").AutoFill Destination:=Range("K7:K" & m)
           
           Sheets("So Phau thuat (2)").Range("J6:K" & m).Copy
           Sheets("So Phau thuat (2)").Range("J6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
           'Sheets("So Phau thuat (2)").Cells(1, 1).Select
        End If
           
        If str = "DuLieu" Then
            strx = str
           
        End If
Next
str = ""

Application.DisplayAlerts = True
Application.ScreenUpdating = True
'Application.Calculation = xlCalculationAutomatic
End Sub

Sub CopyTime2() 'Chua lam xong
Dim i As Variant
Dim str As String
Dim strx As String
Dim m As Integer
Dim a As Integer

Application.DisplayAlerts = False
Application.ScreenUpdating = False
'Application.Calculation = xlCalculationManual

m = Sheets("So Phau thuat").UsedRange.Rows(Sheets("So Phau thuat").UsedRange.Rows.Count).Row

For i = 1 To Sheets.Count
    str = Sheets(i).Name
        If str = "So Phau thuat" Then
            a = 1
        End If
        If str = "DuLieu" Then
           a = 2
        End If
Next
str = ""
If a = 1 Then
    For i = 7 To m
        If Sheets("So Phau thuat").Cells(i, 32) = "Fix" Then
        
        End If
    Next
End If
Application.DisplayAlerts = True
Application.ScreenUpdating = True
'Application.Calculation = xlCalculationAutomatic
End Sub

Sub Changelog()
'22
'Them thu thuat moi: Tap ngoi thang bang tinh va dong
'24
'Sua loi khi tich thu thuat khac thu thuat trong list
'25
'Them tham so m khi them thu thuat moi
End Sub
