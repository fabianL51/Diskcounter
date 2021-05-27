Attribute VB_Name = "CA_CheckBank"
Sub CheckBank()
    
    Set BankDict = CreateObject("Scripting.Dictionary")
    
    Empty1 = 0
    If rng_his.Cells(2, "O").Value = "Bank_Template" Then Empty1 = 1
    Empty2 = 0
    If rng_his.Cells(3, "O").Value = "Bank_Template" Then Empty2 = 1
    Empty3 = 0
    If rng_his.Cells(4, "O").Value = "Bank_Template" Then Empty3 = 1
    
    If Empty1 + Empty2 + Empty3 >= 1 Then
        BankInit.Show
    End If
    Call InitializeBank
End Sub
