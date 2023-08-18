Sub RC_Remove_FCM()
    
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
    Dim dateRow     As Integer
    Dim dateCol     As Integer
    Dim slipCol     As Integer
    Dim rentedNum As Integer
    Dim overlay     As Integer
    Dim underOver As Integer
    Dim skipColor1 As Long 'light green
    Dim skipColor2 As Long 'dark green
    Dim skipColor3 As Long 'teal
    
    
    
    ' Set worksheet to the current active sheet
    Set ws = ActiveSheet
    overlay = 0
    underOver = 0
    ' Set the starting row and the last row to traverse
    startingRow = ActiveCell.Row        ' start from currently selected row
    currentCol = ActiveCell.Column      ' start from currently selected col
    skipColor1 = 15138246
    skipColor2 = 5937414
    skipColor3 = 13433740
    dateCol = FindInRow(1, "Date")
    slipCol = FindInRow(1, "FCM Slips")
    
    While Day(ws.Cells(startingRow, dateCol).Value) <> 1        'if we're not already at the start date, go up
        startingRow = startingRow - 1
    Wend
    
    dateRow = startingRow 'need to save the row that has the first date for later
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row        'i could have just used an infinite loop with an exit, but this works
    
    ' Loop through each row
    For currentRow = startingRow To lastRow
        numSlips = 0
        ' Check if the value in the first column of the current row is a date
        If Not IsDate(ws.Cells(currentRow, dateCol).Value) Then
            Exit For
        End If
        
        If ws.Cells(currentRow, currentCol).Interior.ColorIndex = xlNone Then
            ws.Cells(currentRow, currentCol).Borders(xlEdgeBottom).LineStyle = xlNone
            
            If ws.Cells(currentRow, slipCol).Value = "" Then
                ws.Cells(currentRow, slipCol).Value = 0 'This should never happen
            Else
                ws.Cells(currentRow, slipCol).Value = ws.Cells(currentRow, slipCol).Value - 1
            End If
            
            If (ws.Cells(currentRow - 1, currentCol).Interior.Color <> xlNone) Then
                If ws.Cells(currentRow - 2, currentCol).Interior.Color = xlNone Then
                    ws.Cells(currentRow - 1, currentCol).Borders(xlEdgeBottom).LineStyle = xlNone
                    If ws.Cells(currentRow - 1, slipCol).Value = "" Then
                        ws.Cells(currentRow - 1, slipCol).Value = 0 'This should never happen
                    Else
                        ws.Cells(currentRow - 1, slipCol).Value = ws.Cells(currentRow - 1, slipCol).Value - 1
                    End If
                End If
            End If
        End If
            
    
    ' Move to the next row for the next iteration
    Next currentRow
    
    ws.Cells(currentRow, slipCol).Value = "Under/Over" ' the cell we end on is the right cell for putting the under/over text.
    ws.Cells(currentRow + 1, slipCol).Value = "Overlay"
    If ws.Cells(currentRow + 2, slipCol).Value = "" Then
        ws.Cells(currentRow + 2, slipCol).Value = "rented:"
        ws.Cells(currentRow + 2, slipCol + 1).Value = ws.Cells(dateRow - 3, slipCol + 1).Value
    End If
    
    rentedNum = ws.Cells(currentRow + 2, slipCol + 1).Value
    currentRow = dateRow
    
    While IsNumeric(ws.Cells(currentRow, slipCol))
        ws.Cells(currentRow, slipCol + 1).Value = ws.Cells(currentRow, slipCol).Value - rentedNum 'under over num
        
        If ws.Cells(currentRow, slipCol + 1).Value > 0 Then
            ws.Cells(currentRow, slipCol + 1).Font.Color = vbRed
            overlay = overlay + 1
        End If
        
        underOver = underOver + ws.Cells(currentRow, slipCol + 1).Value
        currentRow = currentRow + 1
    Wend
    
    ws.Cells(currentRow, slipCol + 1).Value = underOver
    ws.Cells(currentRow + 1, slipCol + 1).Value = overlay
    If overlay > 1 Then
        ws.Cells(currentRow + 1, slipCol).Font.Color = vbRed
        ws.Cells(currentRow + 1, slipCol + 1).Font.Color = vbRed
    End If
    
    
End Sub 'Remove_FCM()