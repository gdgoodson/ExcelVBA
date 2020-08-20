Sub fillDeskDuty()
    ' This will fill in the info desk shifts (I) on the FX Daily.
    
    Dim Dsheet As Worksheet ' This is the Daily worksheet. But will probably use active sheet.
    Set Dsheet = ActiveSheet ' Make sure the Daily is the active sheet when running macro.
    Dim staffRow As Variant
    Dim mRow As Variant ' managers row
    Dim staffCol As Variant
    Dim coltotalI As Variant ' how many I are in a column
    Dim rowtotalI As Variant ' how many I are in row
    Dim coltotalV As Variant ' how many VR are in column
    Dim rowtotalV As Variant ' how many VR in row
    
    coltotalI = 0
    rowtotalI = 0
    coltotalV = 0
       
    staffCol = 70 ' code for letter F
    staffRow = 8 ' First row with staff
    RowCount = 1 ' counter
    mRow = 7 ' First row with manager
    totalS = 3
    openRow = 3
    
    managersRows = Array(7, 6, 4, 3)
    staffRowS = Array(8, 9, 10, 11, 12, 13, 14, 15, 16)
    vrRows = Array(18, 19, 20, 5)
    vrTrainedRows = Array(9, 10, 14, 16, 7, 6, 4)
    
    'Entering in open,meeting, closing
    
    'Opening Manager
    opening = False
    While opening = False
        If openRow > 7 Then
            opening = True
        ElseIf Dsheet.Range("E" & openRow).Interior.ColorIndex = -4142 Then
            Dsheet.Range("E" & openRow) = "OPEN"
            opening = True
        Else
            openRow = openRow + 1
        End If
    Wend
    'Closing Manager
    closeRow = 3
    closing = False
    While closing = False
        If closeRow > 7 Then
            closing = True
        ElseIf Dsheet.Range("P" & closeRow).Interior.ColorIndex = -4142 Then
            Dsheet.Range("P" & closeRow) = "CLOSE"
            closing = True
        Else
            closeRow = closeRow + 1
        End If
    Wend
    
    
    'Enter VR shifts here
    For i = 1 To 11
        sCol = Chr(staffCol) ' turns number into letter
        coltotalV = Vcount(Dsheet, sCol)
        If coltolV < 1 Then
            For Each staff In vrRows
                If coltotalV < 1 Then
                    rowtotalV = checkShifts(Dsheet, staff)
                    If rowtotalV < 3 Then
                        If isBlankVR(Dsheet, sCol, staff) Then
                            Dsheet.Range(sCol & staff) = "VR"
                            coltotalV = Vcount(Dsheet, sCol)
                        End If
                    End If
                End If
            Next staff
        End If
        staffCol = staffCol + 1 ' goes to next column
    Next i
    staffCol = 70 ' code for letter F
    
    For i = 1 To 11
        sCol = Chr(staffCol) ' turns number into letter
            coltotalI = Icount(Dsheet, sCol) 'counts I's in row.
            If coltolaI < 2 Then
                While RowCount < 20
                    Randomize
                    staffRow = Int((16 - 8 + 1) * Rnd + 8) ' Choosing a random cell to check. Best way to distribute shifts.
                    ' Checking to see if cell is empty, background color is white, and doesn't have shift preceding it.
                    If isBlank(Dsheet, sCol, staffRow) Then
                        rowtotalI = checkShifts(Dsheet, staffRow) ' Counting I's in staff row
                        If rowtotalI < 3 Then ' If staff have available shifts then
                            Dsheet.Range(sCol & staffRow) = "I"
                            coltotalI = Icount(Dsheet, sCol)
                        End If
                    End If
                    If coltotalI = 2 Then ' If column has two I, then end this.
                        RowCount = 20
                    Else
                        RowCount = RowCount + 1 ' keeping loop going
                    End If
                Wend
                coltotalI = Icount(Dsheet, sCol)
                ' This fills remaining shifts with managers.
                If coltotalI < 2 Then
                    For Each manage In managersRows
                        If coltotalI < 2 Then
                            If isBlank(Dsheet, sCol, manage) Then
                                rowtotalI = checkShifts(Dsheet, manage)
                                If rowtotalI < 2 Then ' If manager's have available shifts then
                                    Dsheet.Range(sCol & manage) = "I"
                                    coltotalI = Icount(Dsheet, sCol)
                                End If
                            End If
                        End If
                        coltotalI = Icount(Dsheet, sCol)
                    Next manage
                End If
            End If
      
        ' resetting counters.
        RowCount = 1
        staffRow = 8
        coltotalI = 0
        staffCol = staffCol + 1 ' goes to next column
    Next i
    

End Sub

Function checkShifts(DailySheet As Worksheet, row As Variant)
checkShifts = WorksheetFunction.CountIf(DailySheet.Range("F" & row & ":P" & row), "I") + WorksheetFunction.CountIf(DailySheet.Range("F" & row & ":P" & row), "VR")

End Function

Function isBlank(DailySheet As Worksheet, col As Variant, row As Variant)
If IsEmpty(DailySheet.Range(col & row)) And DailySheet.Range(col & row).Offset(0, -1).Value <> "I" And DailySheet.Range(col & row).Offset(0, -1).Value <> "VR" And DailySheet.Range(col & row).Interior.ColorIndex = -4142 Then
    isBlank = True
Else
    isBlank = False
End If
End Function

Function isBlankVR(DailySheet As Worksheet, col As Variant, row As Variant)
    If IsEmpty(DailySheet.Range(col & row)) And DailySheet.Range(col & row).Offset(0, -2).Value <> "I" And DailySheet.Range(col & row).Offset(0, -2).Value <> "VR" And DailySheet.Range(col & row).Interior.ColorIndex = -4142 Then
        isBlankVR = True
    Else
        isBlankVR = False
    End If
End Function

Function Icount(DailySheet As Worksheet, col As Variant)
Icount = WorksheetFunction.CountIf(DailySheet.Range(col & "3:" & col & "20"), "I")
End Function

Function Vcount(DailySheet As Worksheet, col As Variant)
Vcount = WorksheetFunction.CountIf(DailySheet.Range(col & "3:" & col & "21"), "VR")
End Function
