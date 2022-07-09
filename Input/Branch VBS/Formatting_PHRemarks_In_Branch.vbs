Sub Formatting_PHRemarks_In_Branch()
'
' Macro8 Macro
'

'
    Range("G1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$K$1:$N$10").AutoFilter Field:=1, Criteria1:=Array("AC", _
        "Grand Total", "LT", "WM-Local", "WM-CBU", "REF"), Operator:=xlFilterValues
        
    Range("G1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("G2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("G1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Selection.End(xlDown).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("G1").Select
    Selection.AutoFilter
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
End Sub
