Attribute VB_Name = "Module1"
'Final Script

'Create a script to loop through all the stocks for the year, for each loop run, extract, and print:
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.

    'Script will also have conditional formatting that will highlight positive change in green and negative change in red.

Sub stocks()

'Define worksheet variable and create loop to iterate through each worksheet
Dim ws As Worksheet
For Each ws In Worksheets
    
    'Create Summary Table headers
    ws.Range("I1").Value = ("ticker")
    ws.Range("J1").Value = ("Yearly" & " " & "Change")
    ws.Range("k1").Value = ("Percent" & " " & "Change")
    ws.Range("l1").Value = ("Total" & " " & "Stock" & " " & "Volume")

'------------------------------------------------------------------
'Calculate ticker tracker and the total volume, percent change and yearly change


    ' Define ticker variable type
    Dim ticker As String
    
    ' Define variable types for opening values
    Dim open_value As Double
    Dim openvalue_ind As Double
    openvalue_ind = 2
    Dim yearlyvalue As Double
    Dim percentchange As Double
    
    'Define variable type for closing value
    Dim close_value As Double
    
    ' Define variable type for holding total per ticker
    Dim tickertotal As Double
    tickertotal = 0
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Loop through all tickers from row 2 until the last row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To lastrow
        open_value = ws.Cells(openvalue_ind, 3).Value
        
    ' Conditional statement to check if still within the same ticker
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
        ' Set the ticker
        ticker = ws.Cells(i, 1).Value
        
        ' Add to the ticker Total
        tickertotal = tickertotal + ws.Cells(i, 7).Value
        
        'Get the closevalue
        close_value = ws.Cells(i, 6).Value
        
        'Calculate yearlychange
        yearlychange = close_value - open_value
        ws.Cells(i, 10).Value = yearlychange

            'Conditional statement for conditional formatting yearlychange
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        
        'calculate percentchange
        If open_value = 0 Then
            percentchange = 0
        Else
            percentchange = yearlychange / open_value
        End If
        
        'Print the values in the Summary Table
        ws.Range("K" & Summary_Table_Row).Value = percentchange
        'Format the percentage in the Summary Table
        ws.Range("K" & Summary_Table_Row) = Format(percentchange, "Percent")
        'Print  yearlychange in the Summary Table
        ws.Range("J" & Summary_Table_Row).Value = yearlychange
        ' Print the ticker in the Summary Table
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ' Print tickertotal to the Summary Table
        ws.Range("L" & Summary_Table_Row).Value = tickertotal
    
        'Conditional statement for conditional formatting the yearlychange
        If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        End If
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset the ticker Total
        tickertotal = 0
        
        ' Reset openvalue_ind to proper number
        openvalue_ind = (i + 1)
        
    ' If the cell immediately following a row is the same ticker
    Else

        ' Add to the ticker Total
        tickertotal = tickertotal + ws.Cells(i, 7).Value
        
    End If
    
  Next i

'------------------------------------------------------------------
' autofit formatting of values

    ws.Range("A:M").Columns.AutoFit

Next ws

End Sub

