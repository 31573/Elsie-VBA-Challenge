Attribute VB_Name = "Module1"
Sub symbol()

'To run on each sheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
    
        'Put the headers for the columns
        Cells(1, "I") = "Ticker"
        Cells(1, "J") = "Open"
        Cells(1, "K") = "Close"
        Cells(1, "L") = "Yearly Change"
        Cells(1, "M") = "Percent Change"
        Cells(1, "N") = "Total Stock Volume"
        Cells(2, "P") = "Greatest % Increase"
        Cells(3, "P") = "Greatest % Decrease"
        Cells(4, "P") = "Greatest Total Volume"
        Cells(1, "Q") = "Ticker_Symbol"
        Cells(1, "R") = "Results"
        
        ' To format headers, max and min cells and columns to bold
        Range(Cells(1, "I"), Cells(1, "R")).Font.Bold = True
        Range(Cells(2, "P"), Cells(4, "P")).Font.Bold = True
        
        ' To format column N to number
        Range("J:L").NumberFormat = "0.00"
        
        ' To format column N to number
        Range("N:N").NumberFormat = "#,###"
        
        'To format column M as percentage
        Range("M:M").NumberFormat = "0.00%"
        
        'To format max and min as percentage
        Range(Cells(2, "R"), Cells(3, "R")).NumberFormat = "0.00%"
    
        'To format Greatest total Volume as number
        Cells(4, "R").NumberFormat = "#,###"
        
        'To set initial variable to hold ticker symbols
        Dim symbol As String
        
        'To set variable for opening price
        Dim openprice As Double
        
        'To set variable fot close price
        Dim closeprice As Double
        
        'To set variable to hold yearly change per ticker symbol
        Dim yrchg As Double
        
        'To set variable for percentage change
        Dim perchg As Double
        
        'To set variable for total Volume
        Dim stockvol As Double
        
        Dim maxAmount As Double
        Dim minAmount As Double
    
        'Keep track of the location for each ticker symbol in the summary table
        Dim itable As Integer
        itable = 2
        
        'Set open price associated with first day of the year
        openprice = Cells(2, "C")
                
        'To set last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop for ticker symbols to the end
        For i = 2 To lastrow
        
            'Check if we are still within the same ticker symbol, if it is not
            If Cells(i + 1, "A") <> Cells(i, "A") Then
            
                'Set the ticker symbol
                symbol = ws.Cells(i, "A")
                                               
                'set closing price associated with last day of the year
                closeprice = Cells(i, "F")
                
                'Set the calculation for the year change
                yrchg = closeprice - openprice
                
                'Set the calculation for the year change
                perchg = yrchg / openprice
                
                'Set the calculation for the total volume per stock
                stockvol = stockvol + Cells(i, "G")
                
                'Print the ticker symbol in the summary table
                Range("I" & itable) = symbol
                
                'Print the open price in the summary table
                Range("J" & itable) = openprice
                
                'Print the close in the summary table
                Range("k" & itable) = closeprice
                
                'Print the year change in the summary table
                Range("L" & itable) = yrchg
                
                'Print the percentage change in the summary table
                Range("M" & itable) = perchg
                
                'Print the totalsticker volume in the summary table
                Range("N" & itable) = stockvol
                                
                'Set condition for value above 0 to be highlighted in green
                If yrchg > 0 Then
                    Cells(itable, "L").Interior.ColorIndex = 4
                    
                'Set condition for value below 0 to be highlighted in red
                ElseIf yrchg < 0 Then
                    Cells(itable, "L").Interior.ColorIndex = 3
                    
                'Set condition  other than above to be white
                Else
                    Cells(itable, "L").Interior.ColorIndex = 2
                End If
                                
                'Add one to the summary table row
                itable = itable + 1
                
                'To house the first open price at the beginning of the year
                openprice = Cells(i + 1, "C")
                
                'Reset the total stock ticker volume
                stockvol = 0
                
                'If the cell in the next row is the same ticker symbol
            Else
                'Set calculation of the stockvolume
                stockvol = stockvol + Cells(i, "G")
                
            End If
        
        Next i
        
        'Set values of Greatest increase and the ticker
        Greatest_increase = 0
        Ticker = ""
        
        'To set last row
        lastrow = Cells(Rows.Count, "I").End(xlUp).Row
        
        'Loop for ticker symbols to the end
        For i = 2 To lastrow
            If Cells(i, "M").Value > Greatest_increase Then
                Greatest_increase = Cells(i, "M")
                Ticker = Cells(i, "I")
            End If
            
        Next i
        
         'Set values of Greatest decrease and the ticker
        Range("Q2") = Ticker
        Range("R2") = Greatest_increase
                   
        Greatest_decrease = 0
        Ticker = ""
        'Loop for ticker symbols to the end
        For i = 2 To lastrow
            If Cells(i, "M").Value < Greatest_decrease Then
                Greatest_decrease = Cells(i, "M")
                Ticker = Cells(i, "I")
            End If
            
        Next i
        
        Range("Q3") = Ticker
        Range("R3") = Greatest_decrease
        
        Greatest_volume = 0
        Ticker = ""
        'Loop for ticker symbols to the end
        For i = 2 To lastrow
            If Cells(i, "N").Value > Greatest_volume Then
                Greatest_volume = Cells(i, "N")
                Ticker = Cells(i, "I")
            End If
            
        Next i
         'Set values of Greatest volume and the ticker
        Range("Q4") = Ticker
        Range("R4") = Greatest_volume
             
             
        'To autofit data in columns
        ws.Cells.Select
        ws.Cells.EntireColumn.AutoFit
        
    'To go to next sheet
    Next ws
    
    ' to return to first sheet
    Sheets("2018").Select
    Range("I1").Select
    
End Sub
    
Sub clear()

'To run on each sheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        'To clear the data
        Range("I:R").clear
        
        ' to return to first sheet
        Sheets("2018").Select
        Range("I1").Select
    
    Next ws
End Sub

