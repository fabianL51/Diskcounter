VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BankInit 
   Caption         =   "UserForm1"
   ClientHeight    =   8364.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18192
   OleObjectBlob   =   "BankInit.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "BankInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False












Public Bank1Okay, Bank2Okay, Bank3Okay As Integer

Private Sub Balance1_Change()
    If Balance1.Enabled = True Then Call ActualizeBankInit
    
End Sub

Private Sub Balance2_Change()
    If Balance2.Enabled = True Then Call ActualizeBankInit
End Sub

Private Sub Balance3_Change()
    If Balance3.Enabled = True Then Call ActualizeBankInit
End Sub

Private Sub Bank1_Click()
    Call ActualizeBankInit
End Sub



Private Sub Bank2_Click()
    ActualizeBankInit
End Sub

Private Sub Bank3_Click()
    Call ActualizeBankInit
End Sub



Private Sub Name1_Change()
    Call ActualizeBankInit
End Sub

Private Sub Name2_Change()
    Call ActualizeBankInit
End Sub

Private Sub Name3_Change()
    Call ActualizeBankInit
End Sub

Private Sub UpdateBank_Click()
    
    Set OldNewBank = CreateObject("Scripting.Dictionary")
    
    If Bank1.Value = True And Bank1Okay = 1 Then
        If rng_his.Cells(2, "O").Value <> Name1.Value And rng_his.Cells(2, "O").Value <> "Bank_Template" Then OldNewBank.Add rng_his.Cells(2, "O").Value, Name1.Value
        rng_his.Cells(2, "O").Value = Name1.Value
        rng_his.Cells(2, "P").Value = CCur(Balance1.Value)
    Else
        rng_his.Cells(2, "O").Value = "Inactive"
        rng_his.Cells(2, "P").Value = 0
    End If
        
    
    If Bank2.Value = True And Bank2Okay = 1 Then
        If rng_his.Cells(3, "O").Value <> Name2.Value And rng_his.Cells(3, "O").Value <> "Bank_Template" Then OldNewBank.Add rng_his.Cells(3, "O").Value, Name2.Value
        rng_his.Cells(3, "O").Value = Name2.Value
        rng_his.Cells(3, "P").Value = CCur(Balance2.Value)
    Else
        rng_his.Cells(3, "O").Value = "Inactive"
        rng_his.Cells(3, "P").Value = 0
    End If

    If Bank3.Value = True And Bank3Okay = 1 Then
        If rng_his.Cells(4, "O").Value <> Name3.Value And rng_his.Cells(4, "O").Value <> "Bank_Template" Then OldNewBank.Add rng_his.Cells(4, "O").Value, Name3.Value
        rng_his.Cells(4, "O").Value = Name3.Value
        rng_his.Cells(4, "P").Value = CCur(Balance3.Value)
    Else
        rng_his.Cells(4, "O").Value = "Inactive"
        rng_his.Cells(4, "P").Value = 0
    End If
    
    BankChanged = 1
    
    Unload Me
    
        
End Sub

Private Sub UserForm_Initialize()
    Balance1.Value = rng_his.Cells(2, "P").Value
    Name1.Value = rng_his.Cells(2, "O").Value
    Balance2.Value = rng_his.Cells(3, "P").Value
    Name2.Value = rng_his.Cells(3, "O").Value
    Balance3.Value = rng_his.Cells(4, "P").Value
    Name3.Value = rng_his.Cells(4, "O").Value
    Call ActualizeBankInit
End Sub


Sub ActualizeBankInit()

    UpdateBank.Enabled = False
    
    
    
    
    If rng_his.Cells(2, "O").Value <> "Inactive" And rng_his.Cells(2, "O").Value <> "Bank_Template" Then
        Bank1.Enabled = False
        Bank1.Value = True
    End If
    
    If Bank1.Value Then
        Name1.Enabled = True
        Balance1.Enabled = True
    Else
        Name1.Enabled = False
        Name1.Value = ""
        Balance1.Value = ""
        Balance1.Enabled = False
    End If
    
    If rng_his.Cells(3, "O").Value <> "Inactive" And rng_his.Cells(2, "O").Value <> "Bank_Template" Then
        Bank2.Enabled = False
        Bank2.Value = True
    End If
    
    If Bank2.Value Then
        Name2.Enabled = True
        Balance2.Enabled = True
    Else
        Name2.Value = ""
        Name2.Enabled = False
        Balance2.Value = ""
        Balance2.Enabled = False
    End If
    
    
    If rng_his.Cells(4, "O").Value <> "Inactive" And rng_his.Cells(2, "O").Value <> "Bank_Template" Then
        Bank3.Enabled = False
        Bank3.Value = True
    End If
    
    If Bank3.Value Then
        Name3.Enabled = True
        Balance3.Enabled = True
    Else
        Name3.Enabled = False
        Name3.Value = ""
        Balance3.Value = ""
        Balance3.Enabled = False
    End If


    
    If Bank1.Value And Not Trim(Name1.Value & vbNullString) = "Bank_Template" And IsNumeric(Balance1.Value) Then
        Bank1Okay = 1
    Else
        Bank1Okay = 0
    End If
    
    If Bank2.Value And Not Trim(Name2.Value & vbNullString) = "Bank_Template" And IsNumeric(Balance2.Value) Then
        Bank2Okay = 1
    Else
        Bank2Okay = 0
    End If
    
    If Bank3.Value And Not Trim(Name3.Value & vbNullString) = "Bank_Template" And IsNumeric(Balance3.Value) Then
        Bank3Okay = 1
    Else
        Bank3Okay = 0
    End If
        
    If Bank1Okay + Bank2Okay + Bank3Okay > 0 Then UpdateBank.Enabled = True
    
    
End Sub

Private Sub UserForm_Terminate()
If UpdateBank.Enabled = False Then End
End Sub
