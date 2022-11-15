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

-[x]    Range() method
    
    
    
End Sub
-[x]    Cells() method
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

- [] total Volume for loop iterator
 totalVolume = 0

Worksheets("2018").Activate
For i = 2 To 3013
    'increase totalVolume

Next i
- [] Sub DQAnalysis()
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
- []Sub DQAnalysis()
    Worksheets("DQ Analysis").Activate

'Make a list of square numbers
For i = 1 To 10

    Cells(1, i).Value = i * i

Next i

End Sub
-[x]     Open a new workbook, insert a module in VBA, create a new macro in the module, and write     the following code into the new macro. Run the macro. What is the value in the cell G1 after     the macro finishes running?
    That’s right! G1 gets filled on the 7th iteration, and 7 squared is 49.
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
- []