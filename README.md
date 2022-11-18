## stock-analysis
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
