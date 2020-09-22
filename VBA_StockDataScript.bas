Attribute VB_Name = "Module1"
Sub StockStats():

'Define all variables

Dim Ticker As String
Dim TickerQuantity As Integer
Dim LastRow As Long
Dim YearOpen As Double
Dim YearClose As Double
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotStockVol As Double
Dim GPI As Double
Dim GPIT As String
Dim GPD As Double
Dim GPDT As String
Dim GSV As Double
Dim GSVT As String

'Set script to run through each worksheet

    For Each ws In Worksheets
    ws.Activate

' Make it neat, add header columns
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
'Locate the last row

    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

'Set values for variables in the worksheet

    Ticker = ""
    TickerQuantity = 0
    YearlyChange = 0
    YearOpen = 0
    PercentChange = 0
    TotStockVol = 0
    
'Account for each cell in the row, excluding the first

    For i = 2 To LastRow

'Find first column for a compiled list of Tickers
    ' Count the tickers

            Ticker = Cells(i, 1).Value
        
    ' Find the opening price for each ticker

         If YearOpen = 0 Then
         YearOpen = Cells(i, 3).Value
         End If

'Find data for the 4th column Total Stock Volume
    'Calculate the total stock volume

            TotStockVol = TotStockVol + Cells(i, 7).Value
        
' Reset the count of the ticker when we reach a new ticker symbol

    If Cells(i + 1, 1).Value <> Ticker Then
    TickerQuantity = TickerQuantity + 1
    Cells(TickerQuantity + 1, 9) = Ticker
            
' Find the closing price of each ticker

    YearClose = Cells(i, 6)
 
'Find data for the second column Yearly Change
    ' Calculate the yearly change

            YearlyChange = YearClose - YearOpen
            
    ' Add yearly change to the worksheet

            Cells(TickerQuantity + 1, 10).Value = YearlyChange
            
    ' If yearly change value is greater than 0, shade cell green

                If YearlyChange > 0 Then
                Cells(TickerQuantity + 1, 10).Interior.ColorIndex = 4
                
    ' If yearly change value is less than 0, shade cell red.

                ElseIf YearlyChange < 0 Then
                Cells(TickerQuantity + 1, 10).Interior.ColorIndex = 3
           

         End If
            
'Find data for the third column Percent Change
    ' Calculate percent change value for ticker
                If YearOpen = 0 Then
                    PercentChange = 0
             Else
                    PercentChange = (YearlyChange / YearOpen)
                End If
            
            
    ' Format the PercentChange value as a percent

            Cells(TickerQuantity + 1, 11).Value = Format(PercentChange, "Percent")
            
         
' Reset the opening price

            YearOpen = 0
            
' Add total stock volume to the worksheet

            Cells(TickerQuantity + 1, 12).Value = TotStockVol
            
' Reset total stock volume

            TotStockVol = 0
        End If
        
    Next i
    
'Second Summary Table for % increase, decrease and Greatest total volume
    'Label columns
    
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
    
    'Locate the last row
    
        LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
        
    ' Define variables and place them in the table
    
        GPI = Cells(2, 11).Value
        GPIT = Cells(2, 9).Value
        GPD = Cells(2, 11).Value
        GPDT = Cells(2, 9).Value
        GSV = Cells(2, 12).Value
        GSVT = Cells(2, 9).Value
    
    
    ' Begin loop
    
        For i = 2 To LastRow
    
    ' Locate ticker with greatest percent increase
    
        If Cells(i, 11).Value > GPI Then
            GPI = Cells(i, 11).Value
            GPIT = Cells(i, 9).Value
        End If
        
    ' Locate ticker with greatest percent decrease
        If Cells(i, 11).Value < GPD Then
            GPD = Cells(i, 11).Value
            GPDT = Cells(i, 9).Value
        End If
        
    ' Locate ticker with greatest stock volume
        If Cells(i, 12).Value > GSV Then
            GSV = Cells(i, 12).Value
            GSVT = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add values for greatest percent increase, decrease, and stock volume to each worksheet
    
    Range("P2").Value = Format(GPIT, "Percent")
    Range("Q2").Value = Format(GPI, "Percent")
    Range("P3").Value = Format(GPDT, "Percent")
    Range("Q3").Value = Format(GPD, "Percent")
    Range("P4").Value = GSVT
    Range("Q4").Value = GSV
    
Next ws


End Sub
