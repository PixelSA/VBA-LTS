  'modified excelintowordandpdf and generate monthly statement for automated charter periods insertion:
    ' excelintowordandpdf takes new parameter (string collection)
    ' generatemonthlystatement new code;
    ' Dim boatName As String replace by sBoatName
    Dim lastRow, Row As Long
    Dim boatCol, startCol, endCol As Integer
    Dim startDate, endDate As Date
    Dim currMonth As Integer
    ' Dim monthStr As String replace by sMonth
    Dim charterPeriods As Collection
    Dim charterString, sPeriodsString As String
    Dim Item As Variant
    Dim itemCount As Long
    Dim wordApp As Object
    Dim wordDoc As Object
    Dim tbl, tblRow As Object
    
    Set charterPeriods = New Collection
    
    currMonth = Month(DateValue(sMonth & " 1"))
    
    Sheets("Charters").Activate
    boatCol = FindInRow(1, "Boat")
    startCol = FindInRow(1, "Start (noon)")
    endCol = FindInRow(1, "End (noon)")
    lastRow = ActiveSheet.UsedRange.Rows.Count
    itemCount = 1
    
    For Row = 1 To lastRow
        If ActiveSheet.Cells(Row, boatCol).Value = sBoatName Then
            If IsDate(ActiveSheet.Cells(Row, startCol)) Then
                startDate = ActiveSheet.Cells(Row, startCol)
            End If
            If IsDate(ActiveSheet.Cells(Row, endCol)) Then
                endDate = ActiveSheet.Cells(Row, endCol)
            End If
            If Month(startDate) = currMonth And Month(endDate) = currMonth Then
                charterString = monthName(Month(startDate)) & " " & CStr(Day(startDate)) & " - " & monthName(Month(endDate)) & " " & CStr(Day(endDate))
                'charterString = monthName(currMonth)
                charterPeriods.Add Item:=charterString
            ElseIf Month(startDate) = currMonth And Month(endDate) <> currMonth Then
                charterString = monthName(Month(startDate)) & " " & CStr(Day(startDate)) & " - " & monthName(Month(endDate)) & " " & CStr(Day(endDate)) & " (pro-rated)"
                charterPeriods.Add Item:=charterString
            ElseIf Month(startDate) <> currMonth And Month(endDate) = currMonth Then
                charterString = monthName(Month(startDate)) & " " & CStr(Day(startDate)) & " - " & monthName(Month(endDate)) & " " & CStr(Day(endDate)) & " (pro-rated)"
                charterPeriods.Add Item:=charterString
            End If
        End If
    Next Row
    sPeriodsString = ""
    For Each Item In charterPeriods
        If itemCount <> charterPeriods.Count Then
            sPeriodsString = sPeriodsString & Space(2) & Item & vbCrLf
        Else
            sPeriodsString = sPeriodsString & Space(2) & Item
        End If
        itemCount = itemCount + 1
    Next Item
    Debug.Print "Full String: " & sPeriodsString
    'Debug.Print ThisWorkbook.Path
    'Set wordApp = CreateObject("Word.Application") 'open word
    'wordApp.Visible = True
    'wordFilePath = ThisWorkbook.Path & "\0 [BOATNAME] Statement for [MONTH].docx" 'grab the statement doc
    'Debug.Print wordFilePath
    'Set wordDoc = wordApp.Documents.Open(wordFilePath)
    'Set tbl = wordDoc.Tables(2)
    'wordApp.Selection.GoTo what:=wdGoToBookmark, Name:="charterPeriods"
    'wordApp.Selection.TypeText Text:=periodsString

' BOAT spreadsheet column has Month-end numbers to poke into the Statement...
'
'    Copy the data from BOAT's month-column into Word's table...
'
    Dim sMonthName As String
    Dim lnCountItems As Long
    Dim periodsInserted As Boolean
    
    'if we've just inserted the charter periods,
    ' we want to go to next cell, but keep lnCountItems the same...
    ' whilst at the same time ensuring that we don't insert into the charter periods row.
    periodsInserted = False
    lnCountItems = 1
    
    For Each wdCell In wdDoc.Tables(2).Columns(2).Cells
    
        If periodsInserted = True Then
            periodsInserted = False
            wdCell.Range.Text = Format(shBoat.Cells(lnCountItems, iCol_BoatMonth).Value, "##,##0.00")         'currency so use "," and 2 decimals
            
            If shBoat.Cells(lnCountItems, iCol_BoatMonth).Value < 0 Then
                wdCell.Range.Font.Color = vbRed                                     'change negative numbers to RED
            End If
        ElseIf IsNumeric(shBoat.Cells(lnCountItems, iCol_BoatMonth).Value) And lnCountItems <> 5 Then
            wdCell.Range.Text = Format(shBoat.Cells(lnCountItems, iCol_BoatMonth).Value, "##,##0.00")         'currency so use "," and 2 decimals
            
            If shBoat.Cells(lnCountItems, iCol_BoatMonth).Value < 0 Then
                wdCell.Range.Font.Color = vbRed                                     'change negative numbers to RED
            End If
        ElseIf lnCountItems = 5 Then
            periodsInserted = True
            If sPeriodsString <> "" Then
                wdApp.Selection.GoTo what:=wdGoToBookmark, Name:="charterPeriods"
                wdApp.Selection.TypeText Text:=sPeriodsString
            End If
        Else
            If lnCountItems = 1 Then
                sMonthName = shBoat.Cells(lnCountItems, iCol_BoatMonth).Value       'needed later too
                wdCell.Range.Text = sMonthName & ", " & Year(Date)    '1st cell: append the year to the full month  <-----------------------------------------------------------------assumes current year?
            Else
                wdCell.Range.Text = shBoat.Cells(lnCountItems, iCol_BoatMonth).Value                 'just text (although all numbers currently)
            End If
        End If
        If periodsInserted = False Then
            lnCountItems = lnCountItems + 1
        End If
    Next wdCell