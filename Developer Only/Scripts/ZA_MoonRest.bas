Attribute VB_Name = "ZA_MoonRest"
Sub Restore_Moonspense()
    
    iRow = 3
    Set MyWB = ThisWorkbook
    While Not IsEmpty(MyWB.Worksheets("Moonspense").Range("A:H").Cells(RowIndex:=iRow, ColumnIndex:="A").Value)
        MyWB.Worksheets("Moonspense").Range("A:H").Cells(RowIndex:=iRow, ColumnIndex:="E").Value = "DUE" 'OPEN ALL STATUS
        iRow = iRow + 1
    Wend
    
    Call MoonSort
    Call UpdateMoonDictionary
    


    'On Error Resume Next
    'Dim Element As Object
    'For Each Element In ActiveWorkbook.VBProject.VBComponents
    '    ActiveWorkbook.VBProject.VBComponents.Remove Element
    'Next

End Sub
