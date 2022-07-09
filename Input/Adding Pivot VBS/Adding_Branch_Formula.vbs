Sub Adding_Branch_Formula()
'
' Macro1 Macro
'

'
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    Range("I1").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Locked = True
    Selection.FormulaHidden = False
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.NumberFormat = "General"
    ActiveCell.FormulaR1C1 = "Branch"
    Range("I2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],6,3)"
    Range("I2").Select
    Selection.Copy
    Range("H2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
End Sub