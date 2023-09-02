Sub Create_Charter()
    'MsgBox "Create_Charter()"
    
    'using current row LEADS sheet, create a new record (row) in the CHARTER sheet
    If ActiveSheet.name <> "Leads" Then
        MsgBox "Must be on worksheet Leads to create a Charter"
        Exit Sub
    End If
    
    Dim dNewStart As Date
    Dim iLeadRow, iCol As Integer
    Dim vValueUnknownType As Variant
    iLeadRow = ActiveCell.Row

    vValueUnknownType = Cells(iLeadRow, FindInRow(constLeadsChartersHeaderRow, "Start (noon)")).Value
    If Not IsDate(vValueUnknownType) Then
        MsgBox "CreateCharter(): Expected a date in column " & iCol
        Exit Sub
    Else
        dNewStart = vValueUnknownType  'safely save the start-date of the new record (from Lead)
    End If
        
    '(could check that the first 6 cols have data, someday)
    'assume starting with a good Lead record...
    ActiveSheet.Rows(iLeadRow).Select   'not sure the is needed??
    ActiveSheet.Rows(iLeadRow).Copy
    
    Sheets("Charters").Activate
  
' find the location in Charters and "insert" the Leads row...
    Dim dCharterRowStart As Date
    Dim lLastDateRow As Long
    Dim sCharterMonth, sLeadMonth As String
    lLastDateRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, FindInRow(constLeadsChartersHeaderRow, "Start (noon)")).End(xlUp).Row
    Dim iCharterRow As Integer
    For iCharterRow = lLastDateRow To 1 Step -1
        vValueUnknownType = Cells(iCharterRow, 1 + eLCHstart).Value
        If IsDate(vValueUnknownType) Then
            dCharterRowStart = vValueUnknownType        'good charter record
            If dNewStart >= dCharterRowStart Then
                'this is the first Charter record that is "too far up" (but now need to look at green-text to find proper month ie. "Dec"
                
                'isolate the Charter's starting month, ie. "Dec"
                sCharterMonth = Format(dCharterRowStart, "Mmm")    'returns "May"
                sLeadMonth = Format(dNewStart, "Mmm")          'returns "May"
                If sLeadMonth <> sCharterMonth Then
                    'gone too far up so go back down finding the matching month (GREEN background) - maybe several packed together
                    Dim iMMMrow As Integer
                    For iMMMrow = (iCharterRow + 1) To lLastDateRow Step 1
                        sCharterMonth = Cells(iMMMrow, 1 + eLCHstart).Value
                        If sLeadMonth = sCharterMonth Then
                            'found the month where the new record belongs (goes immediately after the green )
                            Exit For
                        Else
                            'keep looking (down) for matching month ie. "Dec"
                        End If
                    Next iMMMrow
                    
                    iCharterRow = iMMMrow          'need to update the row index for all the populating code to follow...
                End If
                
                'okay, new record goes immediately after this one
                ActiveSheet.Rows(iCharterRow).Select       'found match (or close)
                iCharterRow = iCharterRow + 1              'this will be the number of our new row...

                ActiveSheet.Rows(iCharterRow).Insert       'creates a new row AND pastes clipboard (row from Leads)
                ActiveSheet.Rows(iCharterRow).Select       '(not req'd but makes new record easier to find on the screen)
                
                Exit For                                   'exit after inserting new row
            End If
        Else
            'must be a label like "Jan" or "FYE22" so skip it
        End If
    Next iCharterRow    'backwards to the top
    'now the Leads row has been pasted in the Charters sheet
    
    'from now on, use only Charter record data (not Leads), loading the Start date...
    vValueUnknownType = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Start (noon)")).Value
    If Not IsDate(vValueUnknownType) Then
        MsgBox "CreateCharter(): New Charter row doesn't have a Start Date"
        Exit Sub
    Else
        dNewStart = vValueUnknownType  'safely save the start-date of the new record (from Lead)
    End If
    
    Dim iDuration As Integer
    iDuration = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Duration")).Value
    If iDuration <= 0 Then
        MsgBox "CreateCharter() Error: Duration must be greater than zero. (" & iDuration & ")"     'avoid divide-by-zero crashes later
        Exit Sub
    End If
    
'setup the Performance Splits (OwnerSplit, OwnerShare, SLTCShare)...
    Dim sBoatOnlyAddr, sDiscountAddr As String
    sBoatOnlyAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Boat Only")).Address(False, False)  'false,false = relative
    sDiscountAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Discount")).Address(False, False)  'false,false = relative
    
    'OwnerSplit
    Dim sBoatName As String
    Dim iFleetRow As Integer
    Dim sBoatType As String 'grab boat type for agreement creation
    
    sBoatName = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Boat")).Value
    
'remvoe for Nic...
'
'    'switch to Fleet sheet to lookup the boat's owner split...
'    Sheets("Fleet").Activate
'
'    Dim lLastFleetRow As Long
    Dim vOwnerSplit As Variant
    vOwnerSplit = 0.7            '70% hard-coded to work for Nic
    
'    lLastFleetRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row      'last boat "(power cat)" ?
'    For iFleetRow = 1 To lLastFleetRow Step 1
'        If Cells(iFleetRow, 1).Value = sBoatName Then
'            vOwnerSplit = Cells(iFleetRow, constFleetOwnerSplitCol).Value   'found it
'            sBoatType = Cells(iFleetRow, FindInRow(1, "Boat Type").Value this should work?
'            Exit For        'done
'        End If
'    Next iFleetRow    'backwards to the top
'    'did loop end okay?
'    If IsEmpty(vOwnerSplit) Then
'        vOwnerSplit = 0             'just to be safe
'    End If
'
'    Sheets("Charters").Activate 'switch back to Charters
    
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "OwnerSplit")).Value = vOwnerSplit     'save the default percent for this boat on this charter record

    'OwnerShare
    Dim cOwnerShare As Currency
    Dim sOwnerSplitAddr, sBrokerFeeAddr, sSleepAboardAddr As String
    sOwnerSplitAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "OwnerSplit")).Address(False, False)  'false,false = relative
    sBrokerFeeAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "BrokerFee")).Address(False, False)  'false,false = relative
    sSleepAboardAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "SA")).Address(False, False)  'false,false = relative
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "OwnerShare")).Formula = "=(" & sBoatOnlyAddr & "+" & sDiscountAddr & ") * " & sOwnerSplitAddr & " + " & sBrokerFeeAddr & " + " & sSleepAboardAddr
    cOwnerShare = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "OwnerShare")).Value       'save this for updating RevByBoatByDay
    
    'SLTCShare
    Dim cSLTCShare As Currency
    Dim sSkipperAddr, sConcessionAddr As String
    sSkipperAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Skipper")).Address(False, False)  'false,false = relative
    sConcessionAddr = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Concession")).Address(False, False)  'false,false = relative
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "SLTCShare")).Formula = "=(" & sBoatOnlyAddr & "+" & sDiscountAddr & ") * (1-" & sOwnerSplitAddr & ") + " & sSkipperAddr & " + " & sConcessionAddr
    cSLTCShare = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "SLTCShare")).Value       'save this for updating RevByBoatByDay
    


'setup payment schedule...
    Dim sTotalAddress, sPercentAddress As String
    Dim iPercent As Integer
    Dim dDue As Date
    
    sTotalAddress = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Total")).Address(False, False)  'false,false = relative
    
    If IsEmpty(Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bPercent")).Value) Then
        'no manual entry from Leads so regular rules apply...
        If iDuration < 365 Then
            '50% Booking and no Secondary
            Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bPercent")).Value = 0.5      'booking 50%
            Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sPercent")).Value = Empty   'secondary empty (0)
        Else
            '25% booking and 25% Secondary
            Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bPercent")).Value = 0.25     'booking 25%
            Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sPercent")).Value = 0.25     'secondary 25%
        End If
    Else
        'Booking percent (from Leads) isnt' empty, so assume bPercent and sPercent are handwritten (and correct!) so carry-on...
    End If
    'Booking...
    sPercentAddress = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bPercent")).Address(False, False)  'false,false = relative
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bAmount")).Formula = "=" & sTotalAddress & " * " & sPercentAddress
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bDue")).Value = Date + 7  '7-days from now (NOT a formula!)
    'Secondary...
    If IsEmpty(Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sPercent")).Value) Then
        Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sAmount")).Formula = Empty
        Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sDue")).Value = Empty
    Else
        sPercentAddress = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sPercent")).Address(False, False)  'false,false = relative
        Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sAmount")).Formula = "=" & sTotalAddress & " * " & sPercentAddress
        Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sDue")).Value = dNewStart - (8 * 30) '8-mths before (NOT a formula!)
    End If
    'Final...
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fPercent")).Value = 0.5                  'final 50%
    sPercentAddress = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fPercent")).Address(False, False)  'false,false = relative
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fAmount")).Formula = "=" & sTotalAddress & " * " & sPercentAddress
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fDue")).Value = dNewStart - 60         '60-days before (NOT a formula!)
    'Security Deposit...
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdAmount")).Value = 3000#
    Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdDue")).Value = dNewStart - 10         '10-days before (NOT a formula!)
    
    'MsgBox "CreateCharter(): Success - Lead converted into a new Charter record (row " & iCharterRow & " as selected)"



' update RevByBoatByDay...

    'must be after Sept 30,2022 (due to RevByBoat format change)...
    Dim dOldFormatEndDate As Date
    dOldFormatEndDate = #9/30/2022#                 'format change date: Sept 30,2022
    If dNewStart <= dOldFormatEndDate Then
        MsgBox "Create Charter: Sheet RevByBoatByDay can only be automatically updated for Charters after Sep30, 2022"
        Exit Sub
    End If

    'get captain before leaving Charter sheet (for cell's Note)...
    Dim vCaptain As String
    vCaptain = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Captain")).Value
    
    Sheets("RevByBoatByDay").Activate

    'make "daily" rates for Owner and SLTC to use in the final formula...
    Dim sOwnerDay, sSltcDay, sRevFormula As String
    
    'sOwnerDay = Format(cOwnerShare / iDuration, "##,###.00")
    'sSltcDay = Format(cSLTCShare / iDuration, "##,###.00")
    sOwnerDay = Format(cOwnerShare / iDuration, "#####.00")
    sSltcDay = Format(cSLTCShare / iDuration, "#####.00")
    sRevFormula = "=" & sOwnerDay & "+" & sSltcDay
    
    Dim iDateCol, iBoatCol As Integer
    
    'store date & boat column from RevByBoat header row
    iBoatCol = FindInRow(constRevByBoatHeaderRow, sBoatName)
    iDateCol = FindInRow(constRevByBoatHeaderRow, "Date")
    
    'find row (date) where this charter starts...
    lLastDateRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, iDateCol).End(xlUp).Row
    Dim iRevRow As Integer
    For iRevRow = lLastDateRow To 1 Step -1
        vValueUnknownType = Cells(iRevRow, iDateCol).Value
        If IsDate(vValueUnknownType) Then
            If dNewStart = vValueUnknownType Then
                'found the date-row in RevByBoat
                Exit For
            End If
        Else
            'must be a label like "Owner" or "Seabbatical" so skip it
        End If
    Next iRevRow    'backwards to the top
    'error check...
    If iRevRow < 2 Then
        MsgBox "Unable to find start date: " & dNewStart & " (so update RevByBoatByDay manually)"
        Exit Sub
    End If
    
    'start adding the formula in every cell for the charter...
    Dim iDaysLeft As Integer
    iDaysLeft = iDuration
    
    Application.EnableEvents = True   'so sheet-change event will calculate the monthly totals for Owner and Seabbatcial
    
    Do While iDaysLeft >= 0
        'note: need to come thru when 0 just to set the cell color on the last day of charter (no dollars)
        vValueUnknownType = Cells(iRevRow, iDateCol).Value
        If IsDate(vValueUnknownType) Then
            If iDaysLeft >= 1 Then
                Cells(iRevRow, iBoatCol).Formula = sRevFormula     'fill cell with owner+sltc formula
                
                If iDaysLeft = iDuration Then
                    'this is 1st cell of charter
                    Cells(iRevRow, iBoatCol).Interior.Color = vbYellow
                    Cells(iRevRow, iBoatCol).Select                         'position in viewable screen
                    If MsgBox("Is YELLOW correct start date?", vbYesNo) = vbNo Then
                        Exit Do     'abort - user said no
                    End If
                    'add Note...
                    Cells(iRevRow, iBoatCol).NoteText vCaptain & " (" & iDuration & ")"   'append "(28)" to captain's name (note syntax: no "=" for assignment)
                End If
            End If
            Cells(iRevRow, iBoatCol).Interior.Color = vbYellow     'set color (note: the final day doesn't have a formula, just a color)
            iDaysLeft = iDaysLeft - 1                       'continue to fill
        Else
            'must be a label like "Owner" or "Seabbatical" so skip it
        End If
    
        iRevRow = iRevRow + 1
    Loop 'iDaysLeft
    
    Application.EnableEvents = False  'done with calc of monthly totals (so shut off again)
    
    
    
'create Charter worksheet "CaptainLastName" in the boat's directory ("FY 2022 Got Rum Charters.xlsm")

    Dim dstWB As Workbook
    Call OpenBoatXlCharter(iCharterRow, dstWB)      'open the boat charter spreadsheet
   
    Dim sNewSheetName As String
    If SetupCaptainCharterSheet(dstWB, "New", dNewStart, sNewSheetName) = vbFalse Then
        MsgBox "CreateCharter() Error: Captain's worksheet could not be setup. (" & vCaptain & ")"
        Exit Sub
    End If

    'populate the boat's new "CaptainLastName" sheet...
    Dim lRow As Long
    ThisWorkbook.Sheets("Charters").Activate   'does this remain with '1 Bookings'? Is this required?
    
    dstWB.Sheets(sNewSheetName).Range("rnC").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Captain")).Value
    dstWB.Sheets(sNewSheetName).Range("rnBN").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Boat")).Value
    dstWB.Sheets(sNewSheetName).Range("rnSD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Start (noon)")).Value
    dstWB.Sheets(sNewSheetName).Range("rnED").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "End (noon)")).Value
    dstWB.Sheets(sNewSheetName).Range("rnD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Duration")).Value
    dstWB.Sheets(sNewSheetName).Range("SplitOwner").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "OwnerSplit")).Value     '(SLTC split is a local formula)
    dstWB.Sheets(sNewSheetName).Range("rnB").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Boat Only")).Value
    Dim cDiscount As Currency
    cDiscount = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Discount")).Value
    dstWB.Sheets(sNewSheetName).Range("rnDisc").Value = cDiscount
    If cDiscount = 0 Then
        lRow = dstWB.Sheets(sNewSheetName).Range("rnDisc").Row
        dstWB.Sheets(sNewSheetName).Rows(lRow).Hidden = True
    End If
    dstWB.Sheets(sNewSheetName).Range("rnBF").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "BrokerFee")).Value            'flag exempt = BrokerFee
    dstWB.Sheets(sNewSheetName).Range("rnSA").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "SA")).Value
    dstWB.Sheets(sNewSheetName).Range("rnCS").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Skipper")).Value
    dstWB.Sheets(sNewSheetName).Range("rnBC").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Concession")).Value
    If dstWB.Sheets(sNewSheetName).Range("rnBC").Value = 0 Then
        lRow = dstWB.Sheets(sNewSheetName).Range("rnBC").Row
        dstWB.Sheets(sNewSheetName).Rows(lRow).Hidden = True
    End If
    dstWB.Sheets(sNewSheetName).Range("rnDW").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "Waiver")).Value
    'Booking...
    dstWB.Sheets(sNewSheetName).Range("rnBDD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bDue")).Value
    dstWB.Sheets(sNewSheetName).Range("rnBDT").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "bAmount")).Value
    'Secondary...
    Dim vDue As Variant
    vDue = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sDue")).Value
    If IsDate(vDue) Then
        dstWB.Sheets(sNewSheetName).Range("rnSPD").Value = vDue
        dstWB.Sheets(sNewSheetName).Range("rnSPT").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sAmount")).Value
    Else
        lRow = dstWB.Sheets(sNewSheetName).Range("rnSPD").Row
        dstWB.Sheets(sNewSheetName).Rows(lRow).Hidden = True
    End If
    'Final...
    dstWB.Sheets(sNewSheetName).Range("rnFPD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fDue")).Value
    dstWB.Sheets(sNewSheetName).Range("rnFPT").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "fAmount")).Value
    'Security...
    dstWB.Sheets(sNewSheetName).Range("rnSDD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdDue")).Value
    dstWB.Sheets(sNewSheetName).Range("rnSDT").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdAmount")).Value
    'default the new sheet's Cancellation section to being hidden...
    dstWB.Sheets(sNewSheetName).Activate    'switch to boat for FindInCol() to HIDE "Cancellation" rows
    lRow = FindInCol(2, "Cancellation:")    '(assumes ActiveSheet)
    dstWB.Sheets(sNewSheetName).Rows(lRow).Hidden = True
    dstWB.Sheets(sNewSheetName).Rows(lRow + 1).Hidden = True
    dstWB.Sheets(sNewSheetName).Rows(lRow + 2).Hidden = True

    'updates: payments rec'd
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
    dstWB.Sheets(sNewSheetName).Range("rnTRD").Value = Cells(iCharterRow, FindInRow(constLeadsChartersHeaderRow, "sdSent")).Value

'create the Agreement and push data fields into it...
    Dim wordFilePath, pdfFilePath As String
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim tbl, tblRow As Object
    Dim iStartRow As Integer
    Dim rowIndex As Long
    Dim tripDuration As Integer
    Dim fiscalYear As Integer
    Dim startDate As Date
    
    
    iStartRow = ActiveCell.Row
    
    Set wordApp = CreateObject("Word.Application") 'open word
    wordApp.Visible = False
    wordFilePath = ThisWorkbook.Path & "\1 Management\Booking Management\Agreement.docx" 'grab the agreement document
    'Debug.Print wordFilePath
    Set wordDoc = wordApp.Documents.Open(wordFilePath)
    Set tbl = wordDoc.Tables(1)
    Set wordDoc = ActiveDocument
    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="boatName" 'input boat name
    wordApp.Selection.TypeText Text:=ActiveSheet.Cells(iStartRow, FindInRow(1, "Boat")).Value
    
    'wordApp.Selection.GoTo what:=wdGoToBookmark, name:="boatType" 'input boat type (commented out for now)
    'wordApp.Selection.TypeText Text:=sBoatType
    
    'wordApp.Selection.GoTo what:=wdGoToBookmark, name:="Name" 'input Name
    'wordApp.Selection.TypeText Text:=ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value
    
    'wordApp.Selection.GoTo what:=wdGoToBookmark, name:="Phone" 'input Phone (?)
    'wordApp.Selection.TypeText Text:=ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value
    
    'wordApp.Selection.GoTo what:=wdGoToBookmark, name:="Name" 'input Address
    'wordApp.Selection.TypeText Text:=ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value
    
    'wordApp.Selection.GoTo what:=wdGoToBookmark, name:="Email" 'input Email
    'wordApp.Selection.TypeText Text:=ActiveSheet.Cells(iStartRow, FindInRow(1, "email")).Value
    
    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="startDate" 'input start date
    wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "Start (noon)")).Value, "mmm d, yyyy")

    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="endDate" 'input end date
    wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "End (noon)")).Value, "mmm d, yyyy")
    
    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="damageWaiver"
    wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Waiver")).Value)
        
    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="charterFee"
    wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Skipper")).Value)

    wordApp.Selection.GoTo what:=wdGoToBookmark, name:="totalAmount"
    wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Total")).Value)
    
    'deal with tables
    tripDuration = ActiveSheet.Cells(iStartRow, FindInRow(1, "Duration")).Value
    'tripDuration = 390 '(hard code to test the else. (it works))
    '         If iDuration < 365 Then
    '         '50% Booking and no Secondary
    '         Else
    '         '25% booking and 25% Secondary
    If tripDuration < 365 Then
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="fiftyDepositDate" 'input deposit due date
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "bDue")).Value, "mmm d, yyyy")

        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="fiftyDepositAmount" 'input deposit amount
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "bAmount")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="fiftyBalanceDate" 'same for balance
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "fDue")).Value, "mmm d, yyyy")

        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="fiftyBalanceAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "fAmount")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="securityDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "sdDue")).Value, "mmm d, yyyy")
        Set tblRow = tbl.Rows(20)
        tblRow.Delete 'delete all three unneeded rows. we need to do this in reverse or else things get wonky
        Set tblRow = tbl.Rows(19)
        tblRow.Delete
        Set tblRow = tbl.Rows(18)
        tblRow.Delete
        If ActiveSheet.Cells(iStartRow, FindInRow(1, "Concession")).Value <> "" Then 'discount
            wordApp.Selection.GoTo what:=wdGoToBookmark, name:="concessionAmount"
            wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Concession")).Value)
        Else
            Set tblRow = tbl.Rows(14)
            tblRow.Delete
        End If
        If ActiveSheet.Cells(iStartRow, FindInRow(1, "Discount")).Value <> "" Then 'discount
            wordApp.Selection.GoTo what:=wdGoToBookmark, name:="discountAmount"
            wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Discount")).Value)
        Else
            Set tblRow = tbl.Rows(13)
            tblRow.Delete
        End If
     Else
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="quarterDepositDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "bDue")).Value, "mmm d, yyyy")

        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="quarterDepositAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "bAmount")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="quarterPaymentDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "fDue")).Value, "mmm d, yyyy")

        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="quarterPaymentAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "fAmount")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="secondBalanceDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "sDue")).Value, "mmm d, yyyy")

        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="secondBalanceAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "sAmount")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="securityDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "sdDue")).Value, "mmm d, yyyy")
        Set tblRow = tbl.Rows(17)
        tblRow.Delete
        Set tblRow = tbl.Rows(16)
        tblRow.Delete
        If ActiveSheet.Cells(iStartRow, FindInRow(1, "Concession")).Value <> "" Then 'discount
            wordApp.Selection.GoTo what:=wdGoToBookmark, name:="concessionAmount"
            wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Concession")).Value)
        Else
            Set tblRow = tbl.Rows(14)
            tblRow.Delete
        End If
        If ActiveSheet.Cells(iStartRow, FindInRow(1, "Discount")).Value <> "" Then 'discount
            wordApp.Selection.GoTo what:=wdGoToBookmark, name:="discountAmount"
            wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Discount")).Value)
        Else
            Set tblRow = tbl.Rows(13)
            tblRow.Delete
        End If
    End If
    
    If ActiveSheet.Cells(iStartRow, FindInRow(1, "SA")).Value > 0 Then 'sleep aboard
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="sleepAboardAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "SA")).Value)
        
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="sleepAboardDate"
        wordApp.Selection.TypeText Text:=Format(ActiveSheet.Cells(iStartRow, FindInRow(1, "Start (noon)")).Value - 1, "mmm d, yyyy")
    End If
    
    If ActiveSheet.Cells(iStartRow, FindInRow(1, "Skipper")).Value > 0 Then 'skipper
        wordApp.Selection.GoTo what:=wdGoToBookmark, name:="chkSkipperAmount"
        wordApp.Selection.TypeText Text:=CStr(ActiveSheet.Cells(iStartRow, FindInRow(1, "Skipper")).Value)
    End If
    
    fiscalYear = FiscalYearOf(ActiveSheet.Cells(iStartRow, FindInRow(1, "Start (noon)")).Value)
    
    wordFilePath = ThisWorkbook.Path & "\" & ActiveSheet.Cells(iStartRow, FindInRow(1, "Boat")).Value & "\Bookings\FY " & fiscalYear & "\" & ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value
    
    If Dir(wordFilePath, vbDirectory) = "" Then 'if filepath doesn't exist, create it. if it does exist then ignore (problematic?)
        MkDir wordFilePath
    End If
    
    
    pdfFilePath = wordFilePath & "\" & ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value & " Agreement.pdf"
    wordFilePath = wordFilePath & "\" & ActiveSheet.Cells(iStartRow, FindInRow(1, "Captain")).Value & " Agreement.docx"
    
    Dim result As VbMsgBoxResult 'is there a better way to loop until someone says yes?
    
    If MsgBox("Create Agreement?", vbYesNo) = vbYes Then
        Set wordDoc = ActiveDocument
        wordApp.Visible = True
        wordDoc.Activate
        result = MsgBox("Agreement is ready to customize in Word." & vbCrLf & "Press yes to create PDF or no to abort PDF.", vbYesNo)
        If result <> vbNo Then
            wordDoc.ExportAsFixedFormat OutputFileName:=pdfFilePath, ExportFormat:=wdExportFormatPDF 'savepdf
        End If
        wordDoc.SaveAs2 wordFilePath 'save docx
    Else
        wordDoc.SaveAs2 wordFilePath
    End If
    
    wordDoc.Close
    
    wordApp.Quit
    
    Set wordDoc = Nothing
    Set wordApp = Nothing
    
    

'Print Preview invoice/receipt/statement and email...
    'next release... Call PrintPreviewEmail(dstWB, sNewSheetName)


    If MsgBox("CreateCharter: Successful (see selection)." & vbCrLf & "Close boat-charter spreadsheet?", vbYesNo) = vbYes Then
        dstWB.Close True        'close-and-save boat charters
    End If
    Set dstWB = Nothing
    
    If MsgBox("Go to invoicing?", vbYesNo) = vbNo Then 'this is kinda awkward, two pop-ups
        Exit Sub
    End If
    
    Sheets("Charters").Activate 'activate Charters
    
    Call InvoiceUpdate(False)
    
    
    
End Sub     'Create_Charter