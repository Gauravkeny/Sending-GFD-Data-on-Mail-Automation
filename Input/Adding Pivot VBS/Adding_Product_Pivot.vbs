Sub Adding_Product_Pivot(Rowscnt)
'
' Macro2 Macro
'

'
    Sheets.Add After:=ActiveSheet
    Sheets("Sheet2").Select
    Sheets("Sheet2").Name = "Pivot"
    Sheets("Sheet1").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Sheet1!R1C1:R" + Rowscnt + "C23", Version:=6).CreatePivotTable TableDestination:= _
        "Pivot!R3C1", TableName:="Product Wise Pivot", DefaultVersion:=6
    
    Sheets("Pivot").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("Product Wise Pivot").PivotFields("Status")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Product Wise Pivot").PivotFields("Type")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("Product Wise Pivot").PivotFields("ProductCode")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Product Wise Pivot").AddDataField ActiveSheet.PivotTables( _
        "Product Wise Pivot").PivotFields("ComplNo"), "Count of ComplNo", xlCount
        
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "Raw Data"
End Sub