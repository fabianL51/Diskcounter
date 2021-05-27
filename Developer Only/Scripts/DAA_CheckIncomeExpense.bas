Attribute VB_Name = "DAA_CheckIncomeExpense"
Sub CheckIncomeExpense(ColumnNow As Integer)
    
    'HISCATVAL UPDATE IN EVERY FOR INCOME/EXPENSE
    
    Set overrange = MyWB.Worksheets(OverviewNow)
    
    Set IncomeDict = CreateObject("Scripting.Dictionary")
    Set ExpenseDict = CreateObject("Scripting.Dictionary")
    
    'INCOME
    overrange.Range("A4:N15").Sort key1:=overrange.Range("N3"), order1:=xlDescending
    
    
    For iRow_Inc = 4 To 15
        IncomeCat = CStr(rng_over.Cells(iRow_Inc, "A").Value)
        If Not HisCatVal.exists(IncomeCat) Then rng_over.Cells(iRow_Inc, ColumnNow).Value = 0
        If rng_over.Cells(iRow_Inc, "N").Value = 0 Then
            overrange.Range("A" & iRow_Inc & ":M" & iRow_Inc).ClearContents
            overrange.Range("A" & iRow_Inc & ":N" & rng_over.Cells(2, "S").Value).Sort key1:=overrange.Range("N4"), order1:=xlDescending
            IncomeCat = CStr(rng_over.Cells(iRow_Inc, "A").Value)
        End If
        If ExpenseDict.exists(ExpenseCat) Then
            IncomeDict(IncomeCat) = iRow_Inc
        Else
            IncomeDict.Add IncomeCat, iRow_Inc
        End If
        If Trim(rng_over.Cells(iRow_Inc + 1, "A").Value & vbNullString) = vbNullString Then Exit For
    Next
    
    rng_over.Cells(2, "R").Value = iRow_Inc
        
    'EXPENSE
    overrange.Range("A17:N" & rng_over.Cells(3, "S").Value).Sort key1:=overrange.Range("N17"), order1:=xlAscending
    
    For iRow_Exp = rng_over.Cells(3, "Q").Value To rng_over.Cells(3, "S").Value
        ExpenseCat = CStr(rng_over.Cells(iRow_Exp, "A").Value)
        If Not HisCatVal.exists(ExpenseCat) Then rng_over.Cells(iRow_Exp, ColumnNow).Value = 0
        If rng_over.Cells(iRow_Exp, "N").Value = 0 Then
            overrange.Range("A" & iRow_Exp & ":M" & iRow_Exp).ClearContents
            overrange.Range("A" & iRow_Exp & ":N" & rng_over.Cells(3, "S").Value).Sort key1:=overrange.Range("N17"), order1:=xlAscending
            ExpenseCat = CStr(rng_over.Cells(iRow_Exp, "A").Value)
        End If
        If ExpenseDict.exists(ExpenseCat) Then
           ExpenseDict(ExpenseCat) = iRow_Exp
        Else
           ExpenseDict.Add ExpenseCat, iRow_Exp
        End If
        
        If Trim(rng_over.Cells(iRow_Exp + 1, "A").Value & vbNullString) = vbNullString Or iRow_Exp + 1 = MaxRowExp Then Exit For
    Next
    
    rng_over.Cells(3, "R").Value = iRow_Exp
    
    MaxRowInc = rng_over.Cells(2, "S").Value
    MaxRowExp = rng_over.Cells(3, "S").Value
    
End Sub


