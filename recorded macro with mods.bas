Sub Macro1()

    Range("B1").Cut Destination:=Range("A6")
    Range("A6").Select
    Range("B6").Select
    Rows("6:6").Select
    Range("B6").Activate
    Range("B6").Select
    Columns("B:B").Select
    Range("B6").Activate
    Columns("B:B").Select
    Selection.Delete Shift:=xlToLeft
    Range("A7").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("D7").Select
    Selection.Insert Shift:=xlToRight
    ActiveCell.Activate
    Range("C7").Select
                  Cells.Find(What:="test", After:=ActiveCell, LookIn:=xlFormulas2, LookAt _
        :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
        False, SearchFormat:=False).Activate
            Rows(ActiveCell.Row).Select
    Selection.ClearContents
Dim last_row, last_col As Long

'Get last row
last_row = Cells(Rows.Count, 3).End(xlUp).Row

'Get last column
last_col = Cells(6, Columns.Count).End(xlToLeft).Column

'Select entire table
Range(Cells(6, 3), Cells(last_row, last_col)).Select
    Selection.Copy
End Sub