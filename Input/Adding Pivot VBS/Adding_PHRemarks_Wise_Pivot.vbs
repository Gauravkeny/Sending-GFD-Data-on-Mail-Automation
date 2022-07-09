Sub Adding_PHRemarks_Wise_Pivot(Rowscnt)
'
' Macro1 Macro
'

'
    Sheets("Raw Data").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Raw Data!R1C1:R" + Rowscnt + "C23", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R3C9", TableName:="PH Remarks Wise Pivot", DefaultVersion:=6
        
    Sheets("Pivot").Select
    Cells(3, 9).Select
    With ActiveSheet.PivotTables("PH Remarks Wise Pivot").PivotFields("Status")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PH Remarks Wise Pivot").PivotFields("Type")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PH Remarks Wise Pivot").PivotFields("ProductCode")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PH Remarks Wise Pivot").PivotFields("PHRemarks")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PH Remarks Wise Pivot").AddDataField ActiveSheet.PivotTables( _
        "PH Remarks Wise Pivot").PivotFields("ComplNo"), "Count of ComplNo", xlCount
End Sub