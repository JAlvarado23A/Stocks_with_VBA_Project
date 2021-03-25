Sub StocksCleaner()

'this will help us loop through all sheet'
For Each ws In Worksheets

    'Greating vars to help compare stock tickers'
    Dim stock1 As String
    Dim stock2 As String
   
   'Declaring vars to hold openning cost, closing cost, yearly change, percent change, and stock volume'
    Dim openCost As Double
    Dim closingCost As Double
    Dim yearChange As Double
    Dim percentChange As Double
    Dim stockVol As Double
    
    Dim amount As Long
    stockVol = 0
    'Set counter to 1 to help me write summary on second row'
    counter = 1

  'Calculate number of rows'
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    

    'Grab initial stock value for calculations later'
    openCost = ws.Cells(2, 3).Value

    'Begin for loop to extrac information needed'
    For i = 2 To lastRow
        'Grab initial tickers to compare'
        stock1 = ws.Cells(i, 1).Value
        stock2 = ws.Cells(i + 1, 1).Value


        
        'If the tickers are not the same calculate yearly change, percentage, and volume'
        'Once print out new caculation starting on I2, J2, K2, and L2'
        If stock2 <> stock1 Then
            counter = counter + 1
            
            'grab the closing cost to calculate summary'
            closingCost = ws.Cells(i, 6).Value


            'calculate yearly change'
            yearChange = closingCost - openCost

            'calculate percent change'
            'Added a safeguard against dividing by zero'
            If openCost <> 0 Then
                percentChange = (yearChange / openCost)
            Else
                percentChange = 0
            End If

            amount = ws.Cells(i, 7).Value
            stockVol = stockVol + amount

            'print summary'
            ws.Cells(counter, "I").Value = stock1
            ws.Cells(counter, "J").Value = yearChange
            ws.Cells(counter, "K").Value = percentChange
            ws.Cells(counter, "L").Value = stockVol
            
            'reset the stock volume for new calculations'
            stockVol = 0

        'Grab initial stock value for calculations later'
        openCost = ws.Cells(i + 1, 3).Value
            
        Else

            'else update all the amounts'
            amount = ws.Cells(i, 7).Value
            stockVol = stockVol + amount

        End If
        
    
    Next i
    
'-----------------------------------------------------------------------
' Copy the I:L headers on sheet 1 to all other sheets
'-----------------------------------------------------------------------
    Dim headers(4) As Variant
    headers(0) = "Ticker"
    headers(1) = "Yearly Change"
    headers(2) = "Percent Change"
    headers(3) = "Total Stock Value"
    
    'Begin for-loop to rename all columns with only the date'
        For j = 0 To 3
            ws.Cells(1, 9 + j).Value = headers(j)

        Next j

'-----------------------------------------------------------------------
'Let us do some formatiing
'-----------------------------------------------------------------------
    'Percent formatting for percentChange
     ws.Columns("K").NumberFormat = "0.00%"
     
     'Color coding yearly change
        lastPercentRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
            
            For i = 2 To lastPercentRow
                
                If ws.Cells(i, 10).Value >= 0 Then
                    ws.Cells(i, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(i, 10).Interior.ColorIndex = 3
                End If
            
            Next i
        
        
     'Autofit column L (volume)
     ws.Columns("I:L").AutoFit

Next ws



End Sub



