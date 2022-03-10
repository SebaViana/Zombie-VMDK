Sub ZombieVMDKs()


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim RVTools As Workbook: Set RVTools = ThisWorkbook

    Sheets("vMetaData").Select
    
    Dim vCenter As String
    vCenter = Range("D2").Value
    
    Dim ext As String
    ext = "..xlsm"
    
    ActiveWorkbook.Worksheets("vHealth").Copy
        
    Dim Zombie As Workbook: Set Zombie = ActiveWorkbook
        
    ActiveSheet.Range("$E:$E").AutoFilter Field:=3, Criteria1:="Zombie"
    
    ActiveSheet.Range("$B:$B").AutoFilter Field:=2, Criteria1:= _
    "Possibly a Zombie vmdk file! Please check."
    

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="]", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    Columns("C:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("B:B").Select
    Selection.TextToColumns Destination:=Range("B1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
        
                        Dim LastRow
Set sht = ActiveSheet
LastRow = sht.UsedRange.Rows(sht.UsedRange.Rows.Count).Row

For i = LastRow To 1 Step -1
If Rows(i).Hidden = True Then Rows(i).EntireRow.Delete
Next

            Columns("A:A").Select
    Selection.Replace What:="[", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
        
            Range("A1:C1").Select
            With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
                Columns("D:G").Select
    Selection.Delete Shift:=xlToLeft
    
        Range("A1").Select
        ActiveCell.FormulaR1C1 = "Datastore"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "VM Name"
        Range("C1").Select
            ActiveCell.FormulaR1C1 = "File"
    Range("A1").Select
    
    Sheets("vHealth").Name = "Sheet1"
    
    Worksheets("Sheet1").Columns("A:C").AutoFit
    
        ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & Format(Now, "yyyy-mm-dd") & " - Zombie VMDKs - " & vCenter & ".xlsx"
        
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub