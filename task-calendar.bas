Sub ChangeDates()

    'Declare variables
    Dim wsSelection As Worksheet
    Dim wsHolidays As Worksheet
    Dim rngHolidays As Range
    Dim rng2024 As Range
    Dim rng2025 As Range
    Dim cell As Range
    Dim checkCell As Range
    Dim foundMatch As Boolean
    Dim compareDate As Date
    Dim dateToAddYear As Date
    Dim lastRowSelection As Long
    Dim lastRowHolidays As Long
    Dim i As Integer

    'Stop screen updating to improve performance
    Application.ScreenUpdating = False

    'Establish references
    Set wsSelection = ActiveWorkbook.Sheets("SELECTION")
    Set wsHolidays = ActiveWorkbook.Sheets("MADRID HOLIDAYS")

    'Count number of tasks
    lastRowSelection = wsSelection.Cells(wsSelection.Rows.Count, 5).End(xlUp).Row
    
    'Count number of holidays
    lastRowHolidays = wsHolidays.Cells(wsHolidays.Rows.Count, 1).End(xlUp).Row

    'Set the range for 2024 dates
    Set rng2024 = wsSelection.Range("E2:E" & lastRowSelection)

    'Set the range for holidays
    Set rngHolidays = wsHolidays.Range("A2:A" & lastRowHolidays)

    'Set the range for 2025 dates
    Set rng2025 = wsSelection.Range("F2")

    'Create loop to add 1 year, check if the new date falls on a holiday and if it matches, subtract one day
    For Each cell In rng2024
        If IsDate(cell.Value) Then
            dateToAddYear = DateAdd("yyyy", 1, cell.Value)

            ' Repeat the verification 3 times
            For i = 1 To 3
                foundMatch = False
                For Each checkCell In rngHolidays
                    If IsDate(checkCell.Value) Then
                        compareDate = checkCell.Value
                        If dateToAddYear = compareDate Then
                            dateToAddYear = DateAdd("d", -1, dateToAddYear)
                            foundMatch = True
                            Exit For
                        End If
                    End If
                Next checkCell

                'Adjust if it falls on Saturday or Sunday
                If Weekday(dateToAddYear) = vbSaturday Then
                    dateToAddYear = DateAdd("d", -1, dateToAddYear)
                ElseIf Weekday(dateToAddYear) = vbSunday Then
                    dateToAddYear = DateAdd("d", -2, dateToAddYear)
                End If
            Next i

            'Write the new date in the 2025 column
            rng2025.Offset(cell.Row - rng2024.Row, 0).Value = dateToAddYear
        End If
    Next cell

    'Re-enable screen updating
    Application.ScreenUpdating = True

    'Notify when done
    MsgBox "Dates updated.", vbInformation

End Sub
