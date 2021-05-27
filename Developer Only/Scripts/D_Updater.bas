Attribute VB_Name = "D_Updater"
Sub Updater()

    Dim iRow_his As Integer
    Dim iRow_cat As Integer
    
    Application.ErrorCheckingOptions.BackgroundChecking = False
    
    iRow_his = rng_his.Cells(Rows.Count, "G").End(xlUp).Row + 1 'START WITH LAST EMPTY 1ST BANK BALANCE (START OF UPDATE9
    iRow_cat = rng_his.Cells(Rows.Count, "M").End(xlUp).Row
    
    'SORT HISTORY
    cutter_his.Range("A3:I" & (rng_his.Cells(Rows.Count, "A").End(xlUp).Row)).Sort key1:=cutter_his.Range("A3"), order1:=xlAscending
    'SORT CATEGORY HISTORY
    cutter_his.Range("L7:M" & (iRow_cat)).Sort key1:=cutter_his.Range("L7"), order1:=xlAscending
    

    In_MonthValue = WorksheetFunction.Sum(cutter_his.Range("B" & MinRowHis & ":B" & rng_his.Cells(2, "M").Value - 1))
    Out_MonthValue = WorksheetFunction.Sum(cutter_his.Range("C" & MinRowHis & ":C" & rng_his.Cells(2, "M").Value - 1))
    Cat_Month = WorksheetFunction.Sum(cutter_his.Range("M" & MinRowCat & ":M" & iRow_cat))
    MonthDifference = Round(Cat_Month - (In_MonthValue - Out_MonthValue), 2)
    
    In_TotValue = WorksheetFunction.Sum(cutter_his.Range("B4" & ":B" & rng_his.Cells(2, "M").Value - 1))
    Out_TotValue = WorksheetFunction.Sum(cutter_his.Range("C4" & ":C" & rng_his.Cells(2, "M").Value - 1))
    Cat_Total = rng_over.Cells(37, "N").Value
    TotDifference = Round(Cat_Total - (In_TotValue - Out_TotValue), 2)
    
    OnlyUpdate = 0
    
    If Reupdate + TotDifference + MonthDifference <> 0 Then
    
        If Reupdate = 1 Or TotDifference <> 0 Then
            iRow_his = 4
        ElseIf MonthDifference <> 0 Then
            iRow_his = MinRowHis
        End If
        
        'SANDWICH THE EXPENSE ADDER
        Call Restore_Moonspense
        Call MonthlyExpenseAdder
        cutter_his.Range("A3:I" & (rng_his.Cells(Rows.Count, "A").End(xlUp).Row)).Sort key1:=cutter_his.Range("A3"), order1:=xlAscending
        Call Restore_Moonspense
        
        iRow_cat = MinRowCat
        cutter_his.Range("L7:M37").ClearContents 'Reset all Categories to zero
        Set HisCatVal = CreateObject("Scripting.Dictionary")
    ElseIf BankChanged = 1 Then
        iRow_his = 4
    Else
        OnlyUpdate = 1
        Call MonthlyExpenseAdder
        If Transfer = 1 Then Call Restore_Moonspense
    End If
    
    
    
    
            
    Do While Not Trim(rng_his.Cells(iRow_his, "A").Value & vbNullString) = vbNullString
        If Month(rng_his.Cells(iRow_his, "A").Value) > Month(rng_his.Cells(iRow_his - 1, "A").Value) Then
            rng_his.Cells(4, "M").Value = iRow_his
            Call ProcessOverview(1, Month(rng_his.Cells(iRow_his - 1, "A").Value))
            Set HisCatVal = CreateObject("Scripting.Dictionary")
        End If
            
        'Change Moon if Found for this month
        Do While MoonPosDict.exists(rng_his.Cells(iRow_his, "D").Value) And OnlyUpdate = 0
            If rng_moon.Cells(MoonPosDict(rng_his.Cells(iRow_his, "D").Value), "E").Value = "PAID" Then
                cutter_his.Range("A" & (iRow_his) & ":I" & (iRow_his)).ClearContents
                cutter_his.Range("A" & iRow_his & ":I" & (rng_his.Cells(Rows.Count, "A").End(xlUp).Row)).Sort key1:=cutter_his.Range("A3"), order1:=xlAscending
            Else
                rng_moon.Cells(MoonPosDict(rng_his.Cells(iRow_his, "D").Value), "E").Value = "PAID"
                Exit Do
            End If
        Loop
        
        If Trim(rng_his.Cells(iRow_his, "A").Value & vbNullString) = vbNullString Then Exit Do
            
        Category = CStr(rng_his.Cells(iRow_his, "E").Value)
        If Not Trim(Category & vbNullString) = vbNullString Then
            If Not HisCatVal.exists(Category) Then
                HisCatVal.Add Category, rng_his.Cells(iRow_his, "B").Value - rng_his.Cells(iRow_his, "C").Value
            Else
                HisCatVal(Category) = HisCatVal(Category) + rng_his.Cells(iRow_his, "B").Value - rng_his.Cells(iRow_his, "C").Value
            End If
        End If
        
        PayBank = CStr(rng_his.Cells(iRow_his, "F").Value)

        If Trim(rng_his.Cells(iRow_his, "G").Value & vbNullString) = vbNullString Or BankChanged = 1 Then
            'UPDATE BANKS' BALANCE
            For Each BankStr In BankDict.keys
                iBankColumn = BankDict(BankStr)
                Select Case BankStr
                    Case PayBank
                        rng_his.Cells(iRow_his, iBankColumn).Formula = "=" & Col_Letter(CDbl(iBankColumn)) & iRow_his - 1 & "+" & "B" & iRow_his & "-" & "C" & iRow_his
                    Case Else
                        If rng_his.Cells(iRow_his, "B").Value = rng_his.Cells(iRow_his, "C").Value Then 'PROCESSING INTERBANKING TRANSFER
                            TrfDetail = rng_his.Cells(iRow_his, "D").Value
                            Position = InStr(1, TrfDetail, BankStr)
                            If Position = 0 Then 'PASSIVE
                                rng_his.Cells(iRow_his, iBankColumn).Formula = "=" & Col_Letter(CDbl(iBankColumn)) & iRow_his - 1
                            ElseIf Position = 1 Then 'SOURCE
                                rng_his.Cells(iRow_his, iBankColumn).Formula = "=" & Col_Letter(CDbl(iBankColumn)) & iRow_his - 1 & "-" & "C" & iRow_his
                            Else 'TARGET
                                rng_his.Cells(iRow_his, iBankColumn).Formula = "=" & Col_Letter(CDbl(iBankColumn)) & iRow_his - 1 & "+" & "B" & iRow_his
                            End If
                        Else
                            rng_his.Cells(iRow_his, iBankColumn).Value = "=" & Col_Letter(CDbl(iBankColumn)) & iRow_his - 1
                        End If
                End Select
                
            Next
        End If
        
        iRow_his = iRow_his + 1
    Loop
    
    rng_his.Cells(2, "M").Value = iRow_his
    
    Dim ChangeMonth As Integer
    If Month(Date) > Month(rng_his.Cells(iRow_his - 1, "A").Value) Then
        ChangeMonth = 1
    ElseIf Month(Date) = Month(rng_his.Cells(iRow_his - 1, "A").Value) Then
        ChangeMonth = 0
    Else
        MsgBox ("Error 2: Inconsistency Date input in row " & iRow_his)
        End
    End If
        
    Call ProcessOverview(ChangeMonth, Month(rng_his.Cells(iRow_his - 1, "A").Value))
    
End Sub
