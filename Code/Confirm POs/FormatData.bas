Attribute VB_Name = "FormatData"
Option Explicit

Sub SortPOConf()
    Dim vRng As Variant
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim PrevSheet As Worksheet


    Set PrevSheet = ActiveSheet
    Sheets("PO Conf").Select

    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Sort by PO creation date, oldest to newest
    With ActiveWorkbook.Worksheets("PO Conf").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("B1"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange Range(Cells(2, 1), Cells(TotalRows, TotalCols))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Color old POs
    'For each cell in column "Created"
    For Each vRng In Range(Cells(2, 2), Cells(TotalRows, 2))
        Select Case CLng(Format(Date, "yymmdd")) - CLng(Format(vRng.Value, "yymmdd"))
                'Highlight POs older than 7 days pink with red text
            Case Is > 7
                With Range(Cells(vRng.Row, 1), Cells(vRng.Row, 5))
                    .Interior.Color = 13551615
                    .Font.Color = -16383844
                End With
                'Highlight POs 3-4 days old yellow with brown text
            Case 3, 4, 5, 6, 7
                With Range(Cells(vRng.Row, 1), Cells(vRng.Row, 5))
                    .Interior.Color = 11534335
                    .Font.Color = -16365673
                End With
        End Select
    Next
    
    PrevSheet.Select
End Sub

Sub Format473()
    Dim PrevSheet As Worksheet
    Set PrevSheet = ActiveSheet

    Sheets("473").Select
    Rows(1).Delete

    PrevSheet.Select
End Sub
