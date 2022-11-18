###stock-analysis
- [x]  Performing analysis to uncover Wall Street trends
- [x]   Macro test message "Hello World" successful
    Dim testMessage As String
    
    
    testMessage = "Hello World!"
    
    
    MsgBox (testMessage)
    
End Sub
- [x] Uploaded green_stocks.xlms saved changes to GitHub
- [x] Cells() method
       Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    
    Cells(3, 1).Value = "Years"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"

- [x]    Range() method
    
    
    
End Sub
- [x]    Cells() method
        Sub DQAnalysis()
        Worksheets("DQ Analysis").Activate

        Range("A1").Value = "DAQO (Ticker: DQ)"


        Worksheets("DQ Analysis").Activate
    
    
        Range("A1").Value = "DAQO (Ticker: DQ)"
    
        'Create a header row
    
        Cells(3, 1).Value = "Years"
    
        Cells(3, 2).Value = "Total Daily Volume"
    
        Cells(3, 3).Value = "Return"

    
    
    
End Sub

- [x] New Macro DQAnalysis
    Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Worksheets("2018").Activate
    For i = 1 To 8
        MsgBox (Cells(1, i))

    Next i

End Sub

- [x] total Volume for loop iterator
 totalVolume = 0

Worksheets("2018").Activate
For i = 2 To 3013
    'increase totalVolume

Next i
- [x] Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    Worksheets("2018").Activate
    rowStart = 2
    rowEnd = 3013
    totalVolume = 0

    For i = rowStart To rowEnd
        'increase totalVolume
        totalVolume = totalVolume + Cells(i, 8).Value

    Next i

End Sub
- [x]   Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

'Make a list of square numbers
For i = 1 To 10

    Cells(1, i).Value = i * i

Next i

End Sub 
- [x]   Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    Worksheets("2018").Activate

    'set initial volume to zero
    totalVolume = 0

    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    End Sub
    
    Sub DQAnalysis()

    'Create a nested for loop that puts a 1 into the cells of all columns A through J, and
    'rows 1 through 10 (cells A1 - A11, B1 - J11,).'
    Range("A1:J1,A1:a10:j10") = 1
    End Sub
    
Sub DQAnalysis()
    '1) Format the output sheet on the "DQ Analysis" worksheet.
    Worksheets("DQ Analysis").Activate
        Range("A1").Value = "All Stocks (2018)"
        'Create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

    '2) Initialize an array of all tickers.
    
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    
    '3a). Initialize variables for thestarting price and ending price.
        
        Dim startingPrice As Single
        Dim endingPrice As Single
    
    '3b) Activate the data worksheet.
        Worksheets("2018").Activate
        
        
    '3c) Find the number of rows to loop over.
    
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        
    '4) Loop through the tickers.
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data.
        Worksheets("2018").Activate
            For j = 2 To RowCount
            
        '5a) Find total volume fo rthe current ticker.
        If Cells(j, 1).Value = ticker Then
        
                totalVolume = totalVolume + Cells(j, 8).Value
            
            End If
            
        '5b) Find starting price for the current ticker.
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                startingPrice = Cells(j, 6).Value
                
            End If
            
        '5c) Find ending price for the current ticker.
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
                endingPrice = Cells(j, 6).Value
                
                        
            End If
        
        Next j
        
        
    '6) Output the data for the current ticker.
        Worksheets("DQ Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
    

End Sub
- []
- []
- []
- []
- []
- []
- []
- []
- []
- []
