Sub UpdateWorksheetsDate()
    ' This will change the name of the sheets of Daily Master
    ' Enter date into Q1 on first sheet (sheet name should be "Saturday")
    ' Short cut to activate Macro: Ctrl+Shift+D
    
    SD = Worksheets(1).Range("Q1").Value ' Takes date of first Saturday in Q1
    WS = 1 ' Worksheet number
    PP = InputBox("Enter Pay Period number ie, 10.") ' Pay Period number
    If Not IsNumeric(PP) Then
        MsgBox "You must enter a number between 1 and 26."
        Exit Sub
    End If
    If PP < 1 Then
        MsgBox "You must enter a number between 1 and 26."
        Exit Sub
    End If
    If PP > 26 Then
        MsgBox "You must enter a number between 1 and 26."
        Exit Sub
    End If
    For i = 1 To 14 ' This is loops through the sheets in Workbook updating names
        Worksheets(WS).name = Format(SD, "dddd m-d")
        SD = SD + 1
        WS = WS + 1
    Next i
    
    Columns("Q").Hidden = True
    ActiveSheet.Shapes("UpdateSheets").Delete
    ActiveSheet.Shapes("CreateNew").Delete
    ' the following saves file as an Excel document to user desktop with new PP name
    
    newdaily = Environ("USERPROFILE") & "\Desktop\[filename]" & PP 'enter filename prefix here
    
    ActiveWorkbook.SaveAs Filename:=newdaily, FileFormat:=xlWorkbookDefault
   
End Sub


