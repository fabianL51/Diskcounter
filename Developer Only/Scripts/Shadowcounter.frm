VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Shadowcounter 
   Caption         =   "Shadowcounter"
   ClientHeight    =   10044
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   17364
   OleObjectBlob   =   "Shadowcounter.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "Shadowcounter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub Category_Change()
Call ActualizeInput
End Sub

Private Sub Detail_Change()
Call ActualizeInput
End Sub

Private Sub EditBank_Click()
Unload Me
BankInit.Show
Call Initializer
Call Updater
End Sub

Private Sub Expense_Click()
Call ActualizeInput
End Sub

Private Sub Help_Click()
    Call Help
End Sub

Private Sub Income_Click()
Call ActualizeInput
End Sub



Private Sub InterBank_Click()
Call ActualizeInput
End Sub

Private Sub Maney_Change()
Call ActualizeInput
End Sub

Private Sub MoonBank_Change()
    Call ActualizeMoonchanger
End Sub

Private Sub MoonCat_Change()
    Call ActualizeMoonchanger
End Sub

Private Sub MoonCost_Change()
    Call ActualizeMoonchanger
End Sub

Private Sub MoonDate_Change()
    Call ActualizeMoonchanger
End Sub

Private Sub MoonExpense_Change()
    Call ActualizeMoonExp
    Call ActualizeMoonchanger
End Sub

Private Sub Moonstrike_Click()
    If Moonstrike.Caption = "Add Expense" Then
        CurrentMoon = rng_moon.Cells(Rows.Count, "A").End(xlUp).Row + 1
        rng_moon.Cells(CurrentMoon, "A").Value = MoonExpense.Value
        rng_moon.Cells(CurrentMoon, "B").Value = MoonCost.Value
        rng_moon.Cells(CurrentMoon, "C").Value = MoonCat.Value
        rng_moon.Cells(CurrentMoon, "D").Value = MoonDate.Value
        rng_moon.Cells(CurrentMoon, "E").Value = "DUE"
        rng_moon.Cells(CurrentMoon, "F").Value = MoonBank.Value
        rng_moon.Cells(1, "K").Value = rng_moon.Cells(1, "K").Value + 1
    ElseIf Moonstrike.Caption = "Update Expense" Then
        iRow_moon = MoonPosDict(MoonExpense.Value)
        rng_moon.Cells(iRow_moon, "A").Value = MoonExpense.Value
        rng_moon.Cells(iRow_moon, "B").Value = MoonCost.Value
        rng_moon.Cells(iRow_moon, "C").Value = MoonCat.Value
        rng_moon.Cells(iRow_moon, "D").Value = MoonDate.Value
        rng_moon.Cells(iRow_moon, "F").Value = MoonBank.Value
    End If
    
    Call UpdateMoonDictionary
    If Day(Date) >= CDbl(MoonDate.Value) Then
        Call MonthlyExpenseAdder
        Call Updater
    End If
    
    MoonExpense.Value = ""
    
    
    
End Sub

Sub ActualizeMoonExp()
    
    
    If MoonPosDict.exists(MoonExpense.Value) Then
        Moonstrike.Caption = "Update Expense"
        iRow_moon = MoonPosDict(MoonExpense.Value)
        MoonCat.Value = rng_moon(iRow_moon, "C").Value
        MoonCost.Value = rng_moon(iRow_moon, "B").Value
        MoonDate.Value = rng_moon(iRow_moon, "D").Value
        MoonStat.Caption = rng_moon(iRow_moon, "E").Value
        MoonBank.Value = rng_moon(iRow_moon, "F").Value
    ElseIf Not Trim(MoonExpense.Value & vbNullString) = vbNullString Then
        Moonstrike.Caption = "Add Expense"
        MoonStat.Caption = "New Expense"
    Else
        Moonstrike.Caption = ""
        MoonCat.Value = ""
        MoonCost.Value = ""
        MoonDate.Value = ""
        MoonBank.Value = ""
    End If
End Sub
Sub ActualizeMoonchanger()
    Moonstrike.Enabled = False
    
    If Not Trim(MoonExpense.Value & vbNullString) = vbNullString Then
        mayday = 0
    Else
        mayday = 1
    End If
    
    If mayday = 0 And Not Trim(MoonCat.Value & vbNullString) = vbNullString Then
        mayday = 0
    Else
        mayday = 1
    End If
    
    If mayday = 0 And IsNumeric(MoonCost.Value) And MoonCost.Value > 0 Then
        mayday = 0
    Else
        mayday = 1
    End If
    
    If mayday = 0 And IsNumeric(MoonDate.Value) And MoonDate.Value > 0 And MoonDate.Value <= 28 Then
        mayday = 0
    Else
        mayday = 1
    End If
    
    If mayday = 0 And BankDict.exists(MoonBank.Value) Then Moonstrike.Enabled = True
        
End Sub

Sub UpdateMoonExp()
    MoonExpense.Clear
    For Each Moonspense In MoonPosDict.keys
        MoonExpense.AddItem Moonspense
    Next
    For Each expensecat In ExpenseDict.keys
        MoonCat.AddItem expensecat
    Next
End Sub



Private Sub PayingBank_Change()
    Call ActualizeInput
End Sub

Private Sub ProcessInput_Click()

    
    Dim iRow_his As Integer
    iRow_his = rng_his.Cells(RowIndex:=2, ColumnIndex:="M").Value
    
    rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="A").Value = Date
    If Income.Value Then
        Code = 2
        rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="B").Value = CCur(Maney.Value)
    ElseIf Expense.Value Then
        Code = 3
        rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="C").Value = CCur(Maney.Value)
    ElseIf InterBank.Value Then
        Code = 7
        rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="C").Value = CCur(Maney.Value)
        rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="B").Value = CCur(Maney.Value)
    End If
    rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="D").Value = Detail.Value
    
    
    If Code <= 3 Then
        rng_his.Cells(RowIndex:=iRow_his, ColumnIndex:="E").Value = Category.Value
        rng_his.Cells(iRow_his, "F").Value = PayingBank.Value
    Else
        rng_his.Cells(iRow_his, "D").Value = SourceBank.Value & " to " & TargetBank.Value
    End If
    
    
    Call Updater
    
    
    Income.Value = False
    Expense.Value = False
    InterBank.Value = False
    Category.Clear
    
End Sub

Sub ActualizeInput()

    ProcessInput.Enabled = False
    
    
    
    If Income.Value Then
        Expense.Enabled = False
        InterBank.Enabled = False
        For Each IncomeCat In IncomeDict.keys
            Category.AddItem IncomeCat
        Next
        Transpage.Value = 0
        FirstStep = 1
    ElseIf Expense.Value Then
        Income.Enabled = False
        InterBank.Enabled = False
        For Each expensecat In ExpenseDict.keys
            Category.AddItem expensecat
        Next
        Transpage.Value = 0
        FirstStep = 1
    ElseIf InterBank.Value Then
        Income.Enabled = False
        Expense.Enabled = False
        For Each expensecat In ExpenseDict.keys
            Category.AddItem expensecat
        Next
        Transpage.Value = 1
        FirstStep = 1
    Else
        Income.Enabled = True
        Expense.Enabled = True
        InterBank.Enabled = True
        Category.Clear
        FirstStep = 0
    End If
    
    MessageUser.BackColor = &HFF00&
    MessageUser.Caption = "Program in Tact"
        
    Select Case Transpage.Value
        Case 0
            If FirstStep = 1 Then
                Category.Enabled = True
            Else
                Category.Value = ""
                Category.Enabled = False
            End If
            
            If Not Trim(Category.Value & vbNullString) = vbNullString Then
                SecondStep = 1
            Else
                SecondStep = 0
            End If
            
            If SecondStep = 1 Then
                PayingBank.Enabled = True
            Else
                PayingBank.Value = ""
                PayingBank.Enabled = False
            End If
                
            If BankDict.exists(PayingBank.Value) Then
                ThirdStep = 1
                Detail.Enabled = True
            Else
                Detail.Enabled = False
                ThirdStep = 0
            End If
                
                
                
        Case 1
            If FirstStep <> 1 Then
                SourceBank.Value = ""
                TargetBank.Value = ""
            End If
            
            If BankDict.exists(SourceBank.Value) And BankDict.exists(TargetBank.Value) Then
                SecondStep = 1
            End If
            
            If SecondStep = 1 And SourceBank.Value <> TargetBank.Value Then
                ThirdStep = 1
                Detail.Enabled = False
            Else
                ThirdStep = 0
                Detail.Enabled = True
            End If
        End Select
            
            If ThirdStep = 1 Then
                Maney.Enabled = True
            Else
                Detail.Text = ""
                Maney.Value = ""
                Maney.Enabled = False
            End If

                
        
    If IsNumeric(Maney.Value) Then ProcessInput.Enabled = True
        
    If Income.Value And IncCatFull = 1 And Not IncomeDict.exists(Category.Value) Then
        ProcessInput.Enabled = False
        MessageUser.BackColor = &HFF&
        MessageUser.Caption = "WARNING: No space left for new income's category"
    ElseIf Expense.Value And ExpCatFull = 1 And Not ExpenseDict.exists(Category.Value) Then
        MessageUser.BackColor = &HFF&
        ProcessInput.Enabled = False
        MessageUser.Caption = "WARNING: No space left for new expense's category"
    End If
    
End Sub



Private Sub SourceBank_Change()
    Call ActualizeInput
End Sub

Private Sub TargetBank_Change()
    Call ActualizeInput
End Sub


Private Sub UserForm_Initialize()
    Call UpdateMoonExp
    Call UpdateBank
End Sub

Sub UpdateBank()
    
    For Each BankName In BankDict.keys
    PayingBank.AddItem BankName
    SourceBank.AddItem BankName
    TargetBank.AddItem BankName
    MoonBank.AddItem BankName
    Next
    
    
End Sub

