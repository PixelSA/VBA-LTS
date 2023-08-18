Sub QAM_FCMslipCalc()

    'confirm user wants to run this before doing anything...
    If MsgBox("Run FCMslipCalc()?", vbYesNo) = vbNo Then
        Exit Sub    'bail out
    End If
    
    MsgBox ("Running FCMslipCalc()... Nic's code goes here...")
    
Dim ws          As Worksheet
    
    ' for looping
    Dim startingRow As Integer
    Dim lastRow     As Integer
    Dim currentRow  As Integer
    Dim currentCol  As Integer
    Dim totalColumns As Integer
    
    ' for calculating slips occupied
    Dim pattern     As Long
    
    ' loop termination
    Dim slipPattern As Long
    
    Dim numSlips    As Integer
    Dim rentedSlips As Integer
    Dim firstCell   As Boolean
    Dim dataRow     As Integer
    Dim dataCol     As Integer
    Dim plusMinusSum As Integer
    Dim overlay     As Integer
    
    ' Set worksheet to the current active sheet
    Set ws = ActiveSheet
    plusMinusSum = 0
    overlay = 0
    ' Set the starting row and the last row to traverse
    startingRow = ActiveCell.Row        ' start from currently selected row for now
    currentCol = FindInRow(1, "Date")
    
    While Day(ws.Cells(startingRow, currentCol).Value) <> 1        'if we're not already at the start date, go up
        startingRow = startingRow - 1
    Wend
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row        'i could have just used an infinite loop with an exit, but this works
    pattern = 17        ' this is the stop pattern as a long
    totalColumns = 40        ' could do infinite loop, but do this instead
    firstCell = False        ' flag for placing data later
    
    ' Loop through each row
    For currentRow = startingRow To lastRow
        numSlips = 0
        ' Check if the value in the first column of the current row is a date
        If Not IsDate(ws.Cells(currentRow, 1).Value) Then
            Exit For        ' Exit the loop if it's not a date
        End If
        
        ' Loop through each column in the current row
        For currentCol = FindInRow(1, "Date") To totalColumns        'starting col is the one with the date header
            If ws.Cells(currentRow, currentCol).Interior.pattern = pattern Then        'if we've reached the end pattern column then stop
                Exit For
            End If
            If ws.Cells(currentRow, currentCol).Borders(xlEdgeBottom).LineStyle = xlDash Then
                numSlips = numSlips + 1        'increment slipcount
            End If
    Next currentCol
    
    ' Move to the next row for the next iteration
    Debug.Print "This row had " & numSlips & "slips. This was row " & currentRow & "."
    currentCol = currentCol + 3        'place us in the right spot to print
    ws.Cells(currentRow, currentCol).Value = numSlips        'print the number of slips
    If firstCell = False Then
        firstCell = True
        dataRow = currentRow
        dataCol = currentCol        'save the initial location we're putting data into for later
    End If
    
Next currentRow

ws.Cells(currentRow, currentCol).Value = "Under/Over"
ws.Cells(currentRow + 1, currentCol).Value = "Overlay"
rentedSlips = ws.Cells(currentRow + 2, currentCol + 1).Value        ' should be where the month's rented slips value is

While IsNumeric(ws.Cells(dataRow, dataCol))
    ws.Cells(dataRow, dataCol + 1).Value = ws.Cells(dataRow, dataCol).Value - rentedSlips
    plusMinusSum = plusMinusSum + (ws.Cells(dataRow, dataCol).Value - rentedSlips)
    
    If (ws.Cells(dataRow, dataCol).Value - rentedSlips) > 0 Then
        overlay = overlay + 1
        ws.Cells(dataRow, dataCol + 1).Font.Color = vbRed
    End If
    dataRow = dataRow + 1
Wend

ws.Cells(dataRow, dataCol + 1).Value = plusMinusSum        'under/over sum
ws.Cells(dataRow + 1, dataCol + 1).Value = overlay        'overlay sum

If overlay <> 0 Then
    ws.Cells(dataRow + 1, dataCol + 1).Font.Color = vbRed
    ws.Cells(dataRow + 1, dataCol).Font.Color = vbRed
End If

    
End Sub 'QAM_FCMslipCalc()