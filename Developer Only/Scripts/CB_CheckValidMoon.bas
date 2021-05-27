Attribute VB_Name = "CB_CheckValidMoon"
Sub CheckValidMoonspense()
    
    'CHECK IF THE EXPENSE HAS VALID BANK ON ITS DISPOSITION
    If MoonPosDict.Count > 0 Then
        For Each Moon In MoonPosDict.items
            Moonspense = CStr(rng_moon(Moon, "A").Value)
            If Not Trim(Moonspense & vbNullString) = vbNullString Then
                MoonBank = CStr(rng_moon(Moon, "F").Value)
                While Not BankDict.exists(MoonBank)
                    Decision = MsgBox("Error 1: The expense " & Moonspense & " has no valid bank for monthly expense", 1)
                    If Decision = 1 Then
                    For Each Bank In BankDict.keys
                    MsgBox (Bank)
                    Next
                        BankInit.Show
                    ElseIf Decision = 0 Then
                        End
                    End If
                Wend
            Else
                Exit For
            End If
        Next
    End If
End Sub
