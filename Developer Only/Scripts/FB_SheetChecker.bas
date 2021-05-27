Attribute VB_Name = "FB_SheetChecker"
Function IsExistSheet(SheetTarget As String) As Boolean
    IsExistSheet = False
    For Each Sheet In ThisWorkbook.Worksheets
        If SheetTarget = Sheet.Name Then
            IsExistSheet = True
            Exit Function
        End If
    Next Sheet
End Function
