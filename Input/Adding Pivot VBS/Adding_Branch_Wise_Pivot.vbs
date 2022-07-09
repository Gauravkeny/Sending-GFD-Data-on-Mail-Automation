Sub Adding_Branch_Wise_Pivot(RowsCnt)
'
' Macro1 Macro
'

'
    Sheets("Raw Data").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Raw Data!R1C1:R" + RowsCnt + "C23", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R3C18", TableName:="Branch Wise Wise Pivot", DefaultVersion:=6
        
    Sheets("Pivot").Select
    Cells(3, 18).Select
    With ActiveSheet.PivotTables("Branch Wise Wise Pivot").PivotFields("Status")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Branch Wise Wise Pivot").PivotFields("Type")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Branch Wise Wise Pivot").PivotFields("Updated Branch")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Branch Wise Wise Pivot").AddDataField ActiveSheet.PivotTables( _
        "Branch Wise Wise Pivot").PivotFields("ComplNo"), "Count of ComplNo", xlCount
End Sub