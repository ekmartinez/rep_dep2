Sub Rep_Dep2()

    Application.ScreenUpdating = False

    Dim lRow As Long
    Dim rcnt As Long

    rcnt = Range("C" & Rows.Count).End(xlUp).Row
    lRow = Range("B" & Rows.Count).End(xlUp).Row

    Worksheets(1).Copy After:=Worksheets(1)
    Worksheets(2).Name = "Data"
    Columns("A:B").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    For I = 1 To rcnt
        If Range("C" & I).Value = "Responsibility Center:" Then
            Range("C" & I).Offset(0, 1).Copy
            Range("A" & I).Offset(3, 0).PasteSpecial xlPasteAll
        End If
        Next I

    For J = 1 To rcnt
        If Range("C" & J).Value = "Account Classification:" Then
            Range("C" & J).Offset(0, 1).Copy
            Range("A" & J).Offset(2, 1).PasteSpecial xlPasteAll
        End If
        Next J
    
    Rows("1:5").Delete
    Columns("F:F").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Range("F2") = "=D2&"" ""&E2"
    Range("F2:F" & lRow).FillDown
    Range("F2:F" & lRow).Copy
    Range("F2:F" & lRow).PasteSpecial xlPasteValues
    Columns("D:E").Delete
    
    Range("A1").Value = "Department"
    Range("B1").Value = "Account"
    Range("C1").Value = "Process Date"
    Range("D1").Value = "Vendor"
    Range("E1").Value = "Expense Description"
    Range("F1").Value = "Invoice Number"
    Range("G1").Value = "Invoice Date"
    Range("H1").Value = "Account Description"
    Range("I1").Value = "Voucher #"
    Range("J1").Value = "Status"
    Range("K1").Value = "Amount"

    Range("$A$1:$K$65000").AutoFilter Field:=11, Criteria1:="="
    Range("$A$2:$K$65000").SpecialCells(xlCellTypeVisible).EntireRow.Delete
    Range("$A$1:$K$65000").AutoFilter

    Columns("A:A").ColumnWidth = 15.14
    Columns("B:B").ColumnWidth = 42.86
    Columns("C:C").ColumnWidth = 14.14
    Columns("D:D").ColumnWidth = 35.71
    Columns("E:E").ColumnWidth = 23.86
    Columns("F:F").ColumnWidth = 14.57
    Columns("G:G").ColumnWidth = 12.57
    Columns("H:H").ColumnWidth = 35.14
    Columns("I:I").ColumnWidth = 8.29
    Columns("J:J").ColumnWidth = 2
    Columns("K:K").ColumnWidth = 11

    With Columns("A:A")
        .HorizontalAlignment = xlLeft
    End With

    With Columns("C:C")
        .NumberFormat = "mm/dd/yyyy"
        .HorizontalAlignment = xlCenter
    End With
    
    With Columns("G:G")
        .NumberFormat = "mm/dd/yyyy"
    End With

    With Columns("E:K")
        .HorizontalAlignment = xlCenter
    End With
    
    Range("A:K").AutoFilter
    
    For Each area In Columns("A:A").SpecialCells(xlCellTypeBlanks)
        If area.Cells.Row <= ActiveSheet.UsedRange.Rows.Count Then
            area.Cells = Range(area.Address).Offset(-1, 0).Value
        End If
        Next area

    For Each area In Columns("B:B").SpecialCells(xlCellTypeBlanks)
        If area.Cells.Row <= ActiveSheet.UsedRange.Rows.Count Then
            area.Cells = Range(area.Address).Offset(-1, 0).Value
        End If
        Next area

End Sub