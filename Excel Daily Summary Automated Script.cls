Sub CopyDataToDailySummaryWithEndingBalanceAndTitle()
    Dim ledgerWorkbook As Workbook
    Dim ledgerSheet As Worksheet
    Dim summarySheet As Worksheet
    Dim currentDate As Date
    
    ' Set the ledger workbook name (without file extension)
    Dim ledgerWorkbookName As String
    ledgerWorkbookName = "Wells Fargo - 2023 GENERAL FUND" ' Change to your ledger workbook's name
    
    ' Check if the ledger workbook is open
    On Error Resume Next
    Set ledgerWorkbook = Workbooks(ledgerWorkbookName)
    On Error GoTo 0
    
    ' If the ledger workbook is not found, display an error message and exit
    If ledgerWorkbook Is Nothing Then
        MsgBox "Workbook '" & ledgerWorkbookName & "' is not open.", vbExclamation
        Exit Sub
    End If
    
    ' Set references to the ledger sheet and summary sheet within the workbook
    Set ledgerSheet = ledgerWorkbook.Sheets("General Fund") ' Change "General Fund" to your actual ledger sheet name
    Set summarySheet = ThisWorkbook.Sheets("Treasurer's Report") ' Change "Treasurer's Report" to your actual daily summary sheet name
    
    ' Get the current date as a serial number (number representing the date)
    currentDate = DateValue(Date)
    
    ' Find the last row in the ledger sheet
    Dim ledgerLastRow As Long
    ledgerLastRow = ledgerSheet.Cells(ledgerSheet.Rows.Count, "A").End(xlUp).Row
    
    Dim ledgerDate As Date
    Dim ledgerRow As Long
    Dim summaryRow As Long
    Dim debitTotal As Double
    Dim creditTotal As Double
    
    ' Clear previous data in the daily summary sheet
    summarySheet.Cells.Clear
    
    ' Add the title at the top of the sheet in large, bold text
    summarySheet.Cells(1, "A").Value = "Treasurer's Report - Concentration Account"
    summarySheet.Cells(1, "A").Font.Size = 18 ' Set the font size
    summarySheet.Cells(1, "A").Font.Bold = True ' Make the text bold
    
    ' Add the current date in cell 2A with the same font size and bold formatting
    summarySheet.Cells(2, "A").Value = currentDate
    summarySheet.Cells(2, "A").Font.Size = 18 ' Set the font size
    summarySheet.Cells(2, "A").Font.Bold = True ' Make the text bold
    
    ' Add the current date and starting balance in C2
    summarySheet.Cells(2, "C").Value = "Starting Balance"
    summarySheet.Cells(2, "C").Font.Bold = True ' Make the text bold
    summarySheet.Cells(2, "D").Value = GetStartingBalance(ledgerSheet, currentDate)
    summarySheet.Cells(2, "D").NumberFormat = "$#,##0.00" ' Format as currency with two decimal places
    summarySheet.Cells(2, "D").Font.Color = RGB(0, 128, 0) ' Set font color to green
    
    ' Adding a double check to the bottom
    summarySheet.Cells(27, "A").Value = "Daily Collections Double Check"
    summarySheet.Cells(27, "A").Font.Size = 14 ' Set the font size
    summarySheet.Cells(27, "A").Font.Bold = True ' Make the text bold
    summarySheet.Cells(27, "A").Font.Italic = True

    ' Set the values for the cells
    summarySheet.Cells(28, "A").Value = "County"
    summarySheet.Cells(29, "A").Value = "Tax Claim"
    summarySheet.Cells(30, "A").Value = "County 2"
    summarySheet.Cells(31, "A").Value = "County 3"
    summarySheet.Cells(32, "A").Value = "Total"
    
    ' Apply underline to the cell containing "County 3"
    summarySheet.Cells(31, "A").Font.Underline = xlUnderlineStyleSingle
    
    ' Format cells in column B as Currency
    summarySheet.Range("B28:B32").NumberFormat = "$#,##0.00"
    
    ' Set up table headers
    summarySheet.Cells(3, "A").Value = "Date"
    summarySheet.Cells(3, "B").Value = "Description"
    summarySheet.Cells(3, "C").Value = "Debit"
    summarySheet.Cells(3, "D").Value = "Credit"
    summarySheet.Cells(3, "E").Value = "Ending Balance"

    ' Add space between the title and the table
    summaryRow = 4 ' Start from row 4 for the table
    
    ' Copy data from the ledger sheet for the current date to the daily summary sheet
    For ledgerRow = 2 To ledgerLastRow
        On Error Resume Next
        ledgerDate = DateValue(ledgerSheet.Cells(ledgerRow, "A").Value) ' Try to convert the ledger date to a date data type
        On Error GoTo 0
        
        If ledgerDate = currentDate Then
            ' Copy date, description, debit, and credit from columns A to D
            summarySheet.Cells(summaryRow, "A").Value = ledgerSheet.Cells(ledgerRow, "A").Value
            summarySheet.Cells(summaryRow, "B").Value = ledgerSheet.Cells(ledgerRow, "B").Value
            summarySheet.Cells(summaryRow, "C").Value = ledgerSheet.Cells(ledgerRow, "C").Value
            summarySheet.Cells(summaryRow, "D").Value = ledgerSheet.Cells(ledgerRow, "D").Value
            
            ' Update the running totals for debits and credits
            debitTotal = debitTotal + summarySheet.Cells(summaryRow, "C").Value
            creditTotal = creditTotal + summarySheet.Cells(summaryRow, "D").Value
            
            summaryRow = summaryRow + 1
        End If
    Next ledgerRow
    
    ' Calculate the ending balance as the sum of starting balance, debits, and credits
    Dim endingBalance As Double
    endingBalance = GetStartingBalance(ledgerSheet, currentDate) + debitTotal - creditTotal
    
    ' Add the ending balance in cell E at the bottom and its red if negative
    summarySheet.Cells(summaryRow + 1, "E").Value = endingBalance
    summarySheet.Cells(summaryRow + 1, "E").NumberFormat = "$#,##0.00" ' Format as currency with two decimal places
    summarySheet.Cells(summaryRow + 1, "E").Font.Bold = True

    ' Check if endingBalance is negative and apply red font color if true
    If endingBalance < 0 Then
        summarySheet.Cells(summaryRow + 1, "E").Font.Color = RGB(255, 0, 0) ' Set font color to red
    End If
    
    ' Add the ending balance description at the bottom
    summarySheet.Cells(summaryRow + 1, "A").Value = "Ending Balance"
    summarySheet.Cells(summaryRow + 1, "A").Font.Bold = True
    
    ' Apply table formatting to the data (optional)
    summarySheet.ListObjects.Add(xlSrcRange, summarySheet.Range("A3:E" & summaryRow + 1), , xlYes).TableStyle = "TableStyleMedium2"
End Sub

Function GetStartingBalance(ledgerSheet As Worksheet, currentDate As Date) As Double
    Dim ledgerLastRow As Long
    Dim ledgerRow As Long
    
    ' Find the last row in the ledger sheet
    ledgerLastRow = ledgerSheet.Cells(ledgerSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Find the starting balance based on the previous date
    For ledgerRow = ledgerLastRow To 2 Step -1
        If DateValue(ledgerSheet.Cells(ledgerRow, "A").Value) < currentDate Then
            GetStartingBalance = ledgerSheet.Cells(ledgerRow, "E").Value
            Exit Function
        End If
    Next ledgerRow
    
    ' Return 0 if no starting balance is found
    GetStartingBalance = 0
End Function


