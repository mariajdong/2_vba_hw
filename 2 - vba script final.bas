Attribute VB_Name = "Module1"
Sub stocks()

'create worksheet variable
Dim sheet As Worksheet

'begin worksheet loop
For Each sheet In Worksheets
    
    'label columns
    sheet.Range("I1") = "ticker"
    sheet.Range("J1") = "yearly change"
    sheet.Range("K1") = "percent change"
    sheet.Range("L1") = "total stock volume"
    
    'bonus rows & columns
    sheet.Range("N2") = "greatest % increase"
    sheet.Range("N3") = "greatest % decrease"
    sheet.Range("N4") = "greatest total volume"
    sheet.Range("O1") = "ticker"
    sheet.Range("P1") = "value"
    
    'calculate last row of data
    Dim lastrow As Long
    lastrow = sheet.Cells(Rows.Count, "A").End(xlUp).Row
    
    'declare variables: opening & closing prices, change in prices, results row, unique tickers
    Dim openprice, closeprice, pricechange, percentchange As Double
    openprice = sheet.Range("C2")
    closeprice = 0
    pricechange = 0
    percentchange = 0

    Dim resultrow As Integer
    resultrow = 2
    
    Dim ticker As String
    ticker = " "
    
    'bonus variables: max/min % change, greatest vol ticker names & values
    Dim maxticker As String
    Dim maxchange As Double
    maxticker = " "
    maxchange = 0
    
    Dim minticker As String
    Dim minchange As Double
    minticker = " "
    minchange = 0
    
    Dim volticker As String
    Dim greatestvol As Double
    volticker = " "
    greatestvol = 0
    
    'begin conditional loop
    For x = 2 To lastrow
        
        'calculate total stock volume
        sheet.Range("L" & resultrow) = sheet.Range("L" & resultrow) + sheet.Cells(x, 7)
        
        'bonus conditional: greatest stock volume
        If sheet.Range("L" & resultrow) > greatestvol Then
            greatestvol = sheet.Range("L" & resultrow)
            volticker = sheet.Cells(x, 1)
        End If
        
        'conditional statement to be used when ticker changes
        If sheet.Cells(x, 1) <> sheet.Cells(x + 1, 1) Then
            closeprice = sheet.Cells(x, 6)
            pricechange = closeprice - openprice
            
            'avoid div by 0 errors, calculate % change
            If openprice <> 0 Then
                percentchange = (pricechange / openprice) * 100
            End If
            
            'populate results
            sheet.Range("I" & resultrow) = sheet.Cells(x, 1)
            sheet.Range("J" & resultrow) = pricechange
            sheet.Range("K" & resultrow) = Str(percentchange) + "%"
            
            'conditional formatting for positive/negative price change
            If pricechange > 0 Then
                sheet.Range("J" & resultrow).Interior.ColorIndex = 4
            ElseIf pricechange <= 0 Then
                sheet.Range("J" & resultrow).Interior.ColorIndex = 3
            End If
            
            'bonus conditional: max/min change
            If percentchange > maxchange Then
                maxchange = percentchange
                maxticker = sheet.Cells(x, 1)
            ElseIf percentchange < minchange Then
                minchange = percentchange
                minticker = sheet.Cells(x, 1)
            End If
            
            'reset stock price variables, move down one row in the results section
            pricechange = 0
            percentchange = 0
            closeprice = 0
            openprice = sheet.Cells(x + 1, 3)
            resultrow = resultrow + 1
        
        End If
    Next x
    
    'bonus values push to sheet
    sheet.Range("O2") = maxticker
    sheet.Range("P2") = Str(maxchange) + "%"
    
    sheet.Range("O3") = minticker
    sheet.Range("P3") = Str(minchange) + "%"
    
    sheet.Range("O4") = volticker
    sheet.Range("P4") = greatestvol
    
Next sheet
    
End Sub
