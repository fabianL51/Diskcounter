Attribute VB_Name = "DA_ProcessOverview"
Sub ProcessOverview(ChangeMonth As Integer, Month As Integer)
    
    iRow_cat = MinRowCat
    For Each Category In HisCatVal.keys
        rng_his.Cells(iRow_cat, "L").Value = Category
        rng_his.Cells(iRow_cat, "M").Value = HisCatVal(Category)
        iRow_cat = iRow_cat + 1
    Next
    
    rng_his.Cells(3, "M").Value = iRow_cat
    
    cutter_his.Range("L" & iRow_cat + 1 & ":M37").ClearContents
    cutter_his.Range("L7:M" & (iRow_cat)).Sort key1:=cutter_his.Range("L7"), order1:=xlAscending
    
    
    
    Dim iColumn As Integer
    iColumn = Month + 1
    
    MaxRowInc = rng_over.Cells(2, "S").Value
    MaxRowExp = rng_over.Cells(3, "S").Value
    
    Call CheckIncomeExpense(iColumn)
    
    Finish_inc = rng_over.Cells(2, "R").Value
    Finish_exp = rng_over.Cells(3, "R").Value
    
    For Each Category In HisCatVal.keys
        CatVal = HisCatVal(Category)
        If CatVal > 0 Then
            If IncomeDict.exists(Category) Then
                iRow_Inc = IncomeDict(Category)
                rng_over.Cells(iRow_Inc, "A").Value = Category
                rng_over.Cells(iRow_Inc, iColumn).Value = CatVal
            Else
                Finish_inc = Finish_inc + 1
                If Finish_inc > MaxRowInc Then
                    MsgBox ("Error 3: Maximum allowed income category exceeded")
                    End
                End If
                rng_over.Cells(Finish_inc, "A").Value = Category
                rng_over.Cells(Finish_inc, iColumn).Value = CatVal
                IncomeDict.Add Category, Finish_inc
            End If
        ElseIf CatVal <= 0 Then
            
            If ExpenseDict.exists(Category) Then
                iRow_Exp = ExpenseDict(Category)
                rng_over.Cells(iRow_Exp, "A").Value = Category
                rng_over.Cells(iRow_Exp, iColumn).Value = CatVal
            Else
                Finish_exp = Finish_exp + 1
                If Finish_exp > MaxRowExp Then
                    overrange.Range("A" & Finish_exp - 1).EntireRow.Insert
                    rng_over.Cells(Finish_exp, "N").Copy rng_over.Cells(Finish_exp - 1, "N")
                    overrange.Range("A" & Finish_exp - 1 & ":N" & Finish_exp).Sort key1:=overrange.Range("N" & Finish_exp - 1), order1:=xlAscending
                End If
                rng_over.Cells(Finish_exp, "A").Value = Category
                rng_over.Cells(Finish_exp, iColumn).Value = CatVal
                ExpenseDict.Add Category, Finish_exp
            End If
        End If
            
    Next
    
    Call CheckIncomeExpense(iColumn)
    

    
    If ChangeMonth = 1 Then
        CalculateSavingPlan (iColumn)
        cutter_his.Range("L7:M37").ClearContents
        For Each Moonspense In MoonPosDict.keys
            iRow_moon = MoonPosDict(Moonspense)
            rng_moon(iRow_moon, "E") = "DUE" 'RESET ALL MONTHLY STATUS
        Next
    End If
    'UPDATE CURRENT BANK BALANCE
    iRow_hislast = rng_his.Cells(2, "M").Value - 1
    
    For iRow_over = 26 To 28
        BankName = CStr(rng_over.Cells(iRow_over, "P").Value)
        If BankDict.exists(BankName) Then
            rng_over.Cells(iRow_over, "R").Value = rng_his.Cells(rng_his.Cells(Rows.Count, "G").End(xlUp).Row, BankDict(BankName)).Value
        End If
    Next
    
    
    Call MoonSort
    Worksheets(HistoryNow).Columns("A:P").AutoFit
    Worksheets("Moonspense").Columns("A:F").AutoFit
    Worksheets(OverviewNow).Columns("A:S").AutoFit
End Sub
