Attribute VB_Name = "Module1"
Sub Stock_Data():



'Assign needed variables
    Dim Ticker As String
    Dim TickerTotal As Integer
    Dim LastRow As Long
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStock As Double
    Dim GPI As Double
    Dim GPIT As String
    Dim GPD As Double
    Dim GPDT As String
    Dim GSV As Double
    Dim GSVT As String

'loop through each worksheet
    For Each ws In Worksheets
    ws.Activate

' Find last row of populated cells
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Add header columns for each worksheet
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("l1").Value = "Total Stock Volume"
    
' set (reset) values for variables
    Ticker = ""
    TickerTotal = 0
    YearOpen = 0
    YearlyChange = 0
    PercentChange = 0
    TotalStock = 0
    
'Set up the loop
For i = 2 To LastRow

' locate the value and assign the variable to the cell value
        Ticker = Cells(i, 1).Value
        
' Locate the opening price and assign the variable
        If YearOpen = 0 Then
            YearOpen = Cells(i, 3).Value
End If
        
' Calculate the stock volume by locating the value within the cells
        TotalStock = TotalStock + Cells(i, 7).Value
        
' Set the loop to restart at each ticker

    If Cells(i + 1, 1).Value <> Ticker Then
    TickerTotal = TickerTotal + 1

    Cells(TickerTotal + 1, 9) = Ticker
            
' Locate the stock closing value

            YearClose = Cells(i, 6)
            
' Calculate yearly change

            YearlyChange = YearClose - YearOpen
            

            Cells(TickerTotal + 1, 10).Value = YearlyChange
            
' If conditional for conditional formatting

    If YearlyChange > 0 Then
    Cells(TickerTotal + 1, 10).Interior.ColorIndex = 4


    ElseIf YearlyChange < 0 Then
    Cells(TickerTotal + 1, 10).Interior.ColorIndex = 3
            
            
' Calculate percent change value
    If YearOpen = 0 Then
    PercentChange = 0
        Else
        PercentChange = (YearlyChange / YearOpen)

End If
            
            
' Format PercentChange cells as a percentage

    Cells(TickerTotal + 1, 11).Value = Format(PercentChange, "Percent")
    
            
           
' Reset

    YearOpen = 0
            
' Find total stock volume

    Cells(TickerTotal + 1, 12).Value = TotalStock
            
' Reset
            TotalStock = 0
End If
End If
Next i
    
    ' Add section to display greatest percent increase, greatest percent decrease, and greatest total volume for each year.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
'Find LastRow

    LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
' Initialize variables and set values of variables initially to the first row in the list.
    GPI = Cells(2, 11).Value
        GPIT = Cells(2, 9).Value
        GPD = Cells(2, 11).Value
        GPDT = Cells(2, 9).Value
        GSV = Cells(2, 12).Value
        GSVT = Cells(2, 9).Value
    
    
' set up the loop
For i = 2 To LastRow
    
' Locate Greatest Percent Increase

    If Cells(i, 11).Value > GPI Then
    GPI = Cells(i, 11).Value

        GPIT = Cells(i, 9).Value

End If
        
'Locate Greatest Percent Decrease

    If Cells(i, 11).Value < GPD Then
    GPD = Cells(i, 11).Value

            GPDT = Cells(i, 9).Value

End If
        
' Locate Greastest Stock Volume

    If Cells(i, 12).Value > GSV Then
        GSV = Cells(i, 12).Value

           GSVT = Cells(i, 9).Value

End If
        
Next i
    
' Assign the variables to a cell in the summary table

    Range("P2").Value = Format(GPIT, "Percent")
    Range("Q2").Value = Format(GPI, "Percent")
    Range("P3").Value = Format(GPDT, "Percent")
    Range("Q3").Value = Format(GPD, "Percent")
    Range("P4").Value = GSVT
    Range("Q4").Value = GSV
    
Next ws
