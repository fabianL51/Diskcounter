Attribute VB_Name = "ZC_Help"
Sub Help()
    Dim WordApp As Object
    Set WordApp = CreateObject("Word.Application")
    'word will be closed while running
    WordApp.Visible = True
    'open the .doc file
    HelpAdress = CStr(Application.ActiveWorkbook.Path) & "\FREEHELP.docx"
    Set WordDoc = WordApp.Documents.Open(HelpAdress)
End Sub

