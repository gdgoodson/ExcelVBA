Sub fillHoursDaily()
    ' To fill in hours of staff from the eschedule
    ' Daily must be active workbook. Eschedules must be open in back ground.
    
    Dim Esheet As Workbook ' This refers to the eSchedules
    Dim Daily As Workbook ' This will be the open daily
    Dim CurPP As Worksheet ' This is the current pay period worksheet
    Dim Dsheet As Worksheet ' This is the Daily worksheet. But will probably use active sheet.
    Dim PP As String  ' Pay Period numger
    Dim staff As String ' Staff's last name.
    Dim DD As Variant ' The date to be entered.
    Dim pcol As Variant ' The column index number of DD in eSchedules.
        
    Set Dsheet = ActiveSheet ' Make sure the Daily is the active sheet when running macro.
    
    PP = InputBox("Enter Pay Period Number, ie, 10.") ' make this an input box for payperiod working with
            ' example: PP#25
    ' PP input error handling
    While Not IsNumeric(PP)
        If Not IsNumeric(PP) Then
            MsgBox "You must enter a number between 1 and 26."
        End If
        PP = InputBox("Enter Pay Period Number, ie, 10.")
    Wend
    While PP > 26 Or PP < 1
        If PP < 1 Then
            MsgBox "You must enter a number between 1 and 26."
        ElseIf PP > 26 Then
            MsgBox "You must enter a number between 1 and 26."
        End If
        PP = InputBox("Enter Pay Period Number, ie, 10.")
    Wend
    PP = "PP#" & PP
    
    DD = InputBox("Enter date: mm/dd/yyyy.") ' make this an input box for the date working with
    While Not IsDate(DD)
        MsgBox "You must enter a valid date."
        DD = InputBox("Enter date: mm/dd/yyyy.")
    Wend
    ppYear = Year(DD)
    DD = DateValue(CDate(DD)) ' Gets numeric value of date.
    Set Esheet = Workbooks(ppYear & "[filename].xlsx") ' Sets the open eSchedules workbook from inputted year
    Set CurPP = Esheet.Worksheets(PP) ' Sets the pay period worksheet
    
    ' Getting the date
    pcol = Application.Match(DD, CurPP.Range("B11:O11").Value, 0) ' This is getting the index number of the date.
    
    sRow = 3 ' The first row in the daily with a staff name.
    For i = 1 To 19 ' Looping through staff in daily. Loop must equal to number of rows looking through.
        ' This gets the last name of staff
        nValue = "B" & sRow ' The cell with the staff name. B is the column that has staff names.
        staff = Dsheet.Range(nValue).Value  ' Getting the staff name in the cell
        first = InStr(staff, ",") ' finding the comma
        tValue = "D" & sRow ' where the time will be entered. D is the column where time is entered.
        If Not first = 0 Then ' Checking to make sure there is a staff name using a comma
            Sname = Mid(staff, 1, first - 1) ' Isolating last name.
            'Bringing it all together with Index/Match. This looks at the eSchedule, matching last name from daily, and fills in time.
            ' The " + 1" is to look at the next row for the time.
            ' Ranges will be dependant on the eSchedule. Another branch or department would need to match with their needs.
                wHours = WorksheetFunction.Index(CurPP.Range("B13:O121"), WorksheetFunction.Match(Sname, CurPP.Range("A13:A121"), 0) + 1, pcol)
                If wHours = "x" Or wHours = "X" Then ' Replacing X's with blanks
                    Dsheet.Range(tValue) = " "
                Else:
                    Dsheet.Range(tValue) = wHours ' Entering the time
                End If
        Else:
            Dsheet.Range(tValue) = " " ' If no staff, enters a blank.
        End If
        sRow = sRow + 1 ' Going to next row.
    Next i
    
    
    ' This will fill cell range in daily and then fill white in cells to represent staff shifts
    
    Dim shift As Variant
    
    Set Dsheet = ActiveSheet ' Make sure the Daily is the active sheet when running macro.
    
    'Available Shifts
    Dim shiftsAll As Variant
    shiftsAll = Array("9:30-6", "12:30-9", "8:45-5:15", "10-2", "10-4", "11:30-5:30", "11:30-6" _
                        , "12-6", "12-5", "1-5", "2-6", "4-9", "5-9", "10-3", "2-8", "9:30-5")
    
    
    ' Start of shift columns
    Set shiftStartC = New Collection
    shiftStartC.Add "E", "9:30-6"
    shiftStartC.Add "E", "9:30-5"
    shiftStartC.Add "I", "12:30-9"
    shiftStartC.Add "E", "8:45-5:15"
    shiftStartC.Add "F", "10-2"
    shiftStartC.Add "F", "10-4"
    shiftStartC.Add "F", "10-3"
    shiftStartC.Add "H", "11:30-5:30"
    shiftStartC.Add "H", "11:30-6"
    shiftStartC.Add "H", "12-5"
    shiftStartC.Add "H", "12-6"
    shiftStartC.Add "I", "1-5"
    shiftStartC.Add "J", "2-6"
    shiftStartC.Add "J", "2-8"
    shiftStartC.Add "L", "4-9"
    shiftStartC.Add "M", "5-9"
    
    ' End of Shift Columns
    Set shiftEndC = New Collection
    shiftEndC.Add "M", "9:30-6"
    shiftEndC.Add "L", "9:30-5"
    shiftEndC.Add "P", "12:30-9"
    shiftEndC.Add "L", "8:45-5:15"
    shiftEndC.Add "I", "10-2"
    shiftEndC.Add "J", "10-3"
    shiftEndC.Add "K", "10-4"
    shiftEndC.Add "L", "11:30-5:30"
    shiftEndC.Add "M", "11:30-6"
    shiftEndC.Add "L", "12-5"
    shiftEndC.Add "M", "12-6"
    shiftEndC.Add "L", "1-5"
    shiftEndC.Add "M", "2-6"
    shiftEndC.Add "O", "2-8"
    shiftEndC.Add "P", "4-9"
    shiftEndC.Add "P", "5-9"
    
    ' User chooses a color for cell background
    
    colorcode = False
        
    While colorcode = False
        UserColor = InputBox("Enter Color Number 1-56 (see macro workbook Color worksheet for index).") ' Make this an input box
        If Not IsNumeric(UserColor) Then
            MsgBox "You must enter a number."
        ElseIf UserColor > 56 Then
            MsgBox "Your number must be less than 56."
        Else
            colorcode = True
        End If
    Wend
    
    Dsheet.Range("E3:P22").Interior.ColorIndex = UserColor
    
    'Start of Shading white for shift
    staffRow = 3
    For i = 1 To 19
        testC = 0
        staffShift = "D" & staffRow
        x = Dsheet.Range(staffShift).Value ' Shift time
        For Each shift In shiftsAll ' this is checking to make sure the shift is included.
            If x = shift Then
                testC = testC + 1
            End If
        Next shift
        
        If testC = 1 Then ' If the shift exists, start shading cells white
            firstC = shiftStartC(x)
            lastC = shiftEndC(x)
        
            Dsheet.Range(firstC & staffRow & ":" & lastC & staffRow).Interior.ColorIndex = 0
            If x = "12:30-9" Then
                Dsheet.Range("H" & staffRow) = "'/"
                Dsheet.Range("M" & staffRow) = "D/"
            End If
        End If
        staffRow = staffRow + 1
    Next i

    
End Sub
