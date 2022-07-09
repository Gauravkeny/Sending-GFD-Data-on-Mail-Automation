Sub Pivot_Filter()
'
' Pivot_Filter Macro
'

'
    Sheets("Pivot").Select

    '    Filtering With Status

    Dim PT As PivotTable
    Dim Pf As PivotField
    Dim Pi As PivotItem
    
    Set PT = ThisWorkbook.Sheets("Pivot").PivotTables("Branch Wise Wise Pivot")
    Set Pf = PT.PivotFields("Status")
    
    For Each Pi In Pf.PivotItems
        If Pi.Name = "Rejected" Then
           Pi.Visible = True
        ElseIf Pi.Name = "Correction Required" Then
           Pi.Visible = True
        Else
           Pi.Visible = False
        End If
    Next Pi
    
    Set PT = Nothing
    Set Pf = Nothing

'   Filtering With Type

    Set PT = ThisWorkbook.Sheets("Pivot").PivotTables("Branch Wise Wise Pivot")
    Set Pf = PT.PivotFields("Type")
    
    For Each Pi In Pf.PivotItems
        If Pi.Name = "Defective" Then
           Pi.Visible = True
        ElseIf Pi.Name = "Mismatch" Then
           Pi.Visible = True
        Else
           Pi.Visible = False
        End If
    Next Pi

End Sub