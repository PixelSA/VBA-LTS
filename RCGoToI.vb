Sub RC_GoTo_Invoicing()
    Dim captainCol As Integer
    Dim boatCol As Integer
    Dim currentRow As Integer
    Dim fullName As String
    Dim nameArr() As String
    Dim lastName As String
    Dim boatName As String
    Dim invoiceWs As Workbook
    
    Dim basePath As String
    Dim fullPath As String
    
    If ActiveSheet.name = "Charters" Then
        captainCol = FindInRow(1, "Captain")
        boatCol = FindInRow(1, "Boat")
        currentRow = ActiveCell.Row
        basePath = Environ("USERPROFILE") & "\Documents\SLTC\9 SLTC - Documents\Seabbatical\Fleet\" 'this should change based on computer/file setup (the first three directories)
    
        fullName = ActiveSheet.Cells(currentRow, captainCol).Value
        nameArr = Split(fullName, " ")
    
        If UBound(nameArr) >= 0 Then 'this is always true unless we made a mistake
            lastName = nameArr(UBound(nameArr))
        End If
    
        boatName = ActiveSheet.Cells(currentRow, boatCol).Value
        fullPath = basePath & boatName & "\Bookings\FY 2023\FY 2023 " & boatName & " Charters.xlsm"
    
        Set invoiceWs = Workbooks.Open(fullPath)
    
        If Not invoiceWs Is Nothing Then
            invoiceWs.Sheets(lastName).Activate
        End If
    ElseIf ActiveSheet.name = "RevByBoatByDay" Then
        boatName = ActiveSheet.Cells(1, ActiveCell.Column)
        fullName = ActiveCell.Value
        nameArr = Split(fullName, " ")
        
        basePath = Environ("USERPROFILE") & "\Documents\SLTC\9 SLTC - Documents\Seabbatical\Fleet\" 'this should change based on computer/file setup (the first three directories)
    
        If UBound(nameArr) >= 0 Then 'this is always true unless we made a mistake
            lastName = nameArr(UBound(nameArr))
        End If
        
        fullPath = basePath & boatName & "\Bookings\FY 2023\FY 2023 " & boatName & " Charters.xlsm"

        Set invoiceWs = Workbooks.Open(fullPath)
    
        If Not invoiceWs Is Nothing Then
            invoiceWs.Sheets(lastName).Activate
        End If
    End If
End Sub