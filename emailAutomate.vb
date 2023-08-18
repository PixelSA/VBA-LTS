Sub InvoiceUpdate(RC_Called As Boolean)
    
    'using current row CHARTERS sheet, update the boat-charter worksheet for the selected charter sheet entry
    If ActiveSheet.name <> "Charters" Then
        MsgBox "Must be on worksheet Charters to create a Charter Update"
        Exit Sub
    End If
    Dim captainCol As Integer
    Dim fullName As String
    Dim nameArr() As String
    Dim lastName As String
    Dim pdfFilePath As String
    Dim recipientEmail As String
    
    recipientEmail = ActiveSheet.Cells(ActiveCell.Row, FindInRow(1, "email")).Value
        
    
    Dim dNewStart As Date
    Dim iCharterRow As Integer
    Dim iCol As Integer
    Dim vValueUnknownType As Variant
    iCharterRow = ActiveCell.Row
    iCol = ActiveCell.Column
    
    Dim dv As Validation
    Dim vaSplit As Variant
    Dim i As Long
    Dim invoiceDone As Boolean
    Dim outlookApp As Object
    Dim outlookMail As Object
    
    captainCol = FindInRow(1, "Captain")
    fullName = ActiveSheet.Cells(iCharterRow, captainCol).Value
    nameArr = Split(fullName, " ")
    
    If UBound(nameArr) >= 0 Then 'this is always true unless we made a mistake
        lastName = nameArr(UBound(nameArr))
    End If

    vValueUnknownType = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Start (noon)")).Value
    If Not IsDate(vValueUnknownType) Then
        MsgBox "CharterUpdate(): Expected a date in row:" & iCharterRow & " col:" & iCol
        Exit Sub
    Else
        dNewStart = vValueUnknownType  'safely save the start-date of the new record (from Lead)
    End If
        
    '(could check that the first 6 cols have data, someday)
    'assume starting with a good Charters record...
    Dim sCaptain As String
    sCaptain = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Captain")).Value

    Dim dstWB As Workbook
    Call OpenBoatXlCharter(iCharterRow, dstWB)      '-----------------open the boat charter spreadsheet (that becomes Active sheet now too)
    
    Dim sNewSheetName As String
    If SetupCaptainCharterSheet(dstWB, "Find", dNewStart, sNewSheetName) = vbFalse Then
        MsgBox "CharterUpdate() Error: Captain's worksheet could not be setup. (" & sCaptain & ")"
        Exit Sub
    End If
  
    'update the boat's "CaptainLastName" sheet...
    ThisWorkbook.Sheets("Charters").Activate   'does this remain with '1 Bookings'? Is this required?

    'Booking...
    dstWB.Sheets(sNewSheetName).Range("rnBDA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bActual")).Value
    dstWB.Sheets(sNewSheetName).Range("rnBDR").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bOn")).Value
    'Secondary...
    dstWB.Sheets(sNewSheetName).Range("rnSPA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sActual")).Value
    dstWB.Sheets(sNewSheetName).Range("rnSPR").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sOn")).Value
    'Final...
    dstWB.Sheets(sNewSheetName).Range("rnFPA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fActual")).Value
    dstWB.Sheets(sNewSheetName).Range("rnFPR").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fOn")).Value
    'Security...
    dstWB.Sheets(sNewSheetName).Range("rnSDA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdActual")).Value
    dstWB.Sheets(sNewSheetName).Range("rnSDR").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdOn")).Value
    'Refund...
    'dstWB.Sheets(sNewSheetName).Range("rnTRA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdRefund")).Value    'don't overwrite that formula!
    'dstWB.Sheets(sNewSheetName).Range("rnTRD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdSent")).Value
    Dim vRefundDate As Variant
    vRefundDate = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdSent")).Value
    If IsDate(vRefundDate) Then
        dstWB.Sheets(sNewSheetName).Range("rnTRD").Value = vRefundDate  'store Charters date into boat's sheet
        'boat-sheet has formula to calculate final refund (net Waiver), so pick that out and store back onto Charters sheet...
        Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdRefund")).Value = dstWB.Sheets(sNewSheetName).Range("rnTRA").Value
        
        dstWB.Sheets(sNewSheetName).Tab.Color = vbGreen                'flag sheet's tab as finished
    End If
        
        
    'next release... Call PrintPreviewEmail(dstWB, sNewSheetName)
    
    'If MsgBox("Update successful." & vbCrLf & "Close boat-charter spreadsheet?", vbYesNo) = vbYes Then 'should we get rid of this?
        'dstWB.Close True        'close-and-save boat charters
    'End If
    
    dstWB.Sheets(lastName).Activate
    
    If RC_Called = True Then 'increment the drop-down selection *if called from RC menu*
        Set dv = ActiveSheet.Range("C1").Validation
    
        vaSplit = Range(dv.Formula1).Value
    
        For i = LBound(vaSplit, 1) To UBound(vaSplit, 1)
            If vaSplit(i, 1) = ActiveCell.Value Then
                If i < UBound(vaSplit, 1) Then
                    ActiveCell.Value = vaSplit(i + 1, 1)
                    Exit For
                Else
                    invoiceDone = True 'for later
                End If
            End If
        Next i
    End If
    
    'Print Preview invoice/receipt/statement and email...
    pdfFilePath = Environ$("TEMP") & "\" & "tempPrintPDF.pdf"
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If Err.Number <> 0 Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    If Not outlookApp Is Nothing Then
        ' Create a new email
        Set outlookMail = outlookApp.CreateItem(0)
        
        ' Set email properties
        With outlookMail
            .Subject = "Payment " & Range("C1").Value & " For Seabattical"
            .Body = "Please find the PDF attachment of your payment schedule."
            .Attachments.Add pdfFilePath
            
            ' Recipient's email address
            .To = recipientEmail ' Change to the recipient's email address
            
            ' Display email before sending.
            .Display
        End With
        
        ' Release objects
        Set outlookMail = Nothing
        Set outlookApp = Nothing
    Else
        MsgBox "Outlook application could not be found or created."
    End If
    
    ' Delete the temporary PDF file
    Kill pdfFilePath
    
    Set dstWB = Nothing

End Sub 'CharterUpdate