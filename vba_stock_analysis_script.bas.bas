Attribute VB_Name = "Module1"
Sub Main()
    Dim ws As Worksheet
    ' Loop through each worksheet in the workbook
    For Each ws In Worksheets
    ws.Activate
    Debug.Print ws.Name
        ' Format the current worksheet
        FormatAllSheets ws
        ' Process the stocks in the current worksheet
        LoopThruStocks ws
    Next ws
End Sub

Sub FormatAllSheets(ws As Worksheet)
    Dim rng As Range
    Dim LastRow As Long
    Dim LastColumn As Long
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Determine the last row and column in the current worksheet
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        ' Define the range of columns you want to format as numbers
        Set rng = ws.Range(ws.Cells(1, 3), ws.Cells(LastRow, 6))
        
        ' Format the range as numbers with two decimal places
        rng.NumberFormat = "0.00"
    Next ws
End Sub


Sub LoopThruStocks(ws As Worksheet)
    ' Declare variables
    Dim i As Long
    Dim Stock_Name As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Quarterly_Change As Double
    Dim Percentage_Change As Double
    Dim Stock_Total As Double
    
    'Variables to track greatest % increase, % decrease, and max volume
    Dim Max_Percent_Increase As Double
    Dim Min_Percent_Decrease As Double
    Dim Max_Volume As Double
    
    Dim Max_Percent_Stock As String
    Dim Min_Percent_Stock As String
    Dim Max_Volume_Stock As String
    
    'Initialize variables
    Stock_Total = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    Max_Percent_Increase = 0
    Min_Percent_Decrease = 0
    Max_Volume = 0
    Max_Volume_Stock = " "
    Min_Percent_Stock = " "
    
    ' Determine the last row of data
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Loop through all the stock amounts
    For i = 2 To LastRow
        
        ' Check if we are still within the same Stock Ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set Stock Name
            Stock_Name = Cells(i, 1).Value
            
            ' Find the first row of the current Stock Ticker
            Dim FirstTickerRow As Long
            FirstTickerRow = i
            Do While Cells(FirstTickerRow - 1, 1).Value = Stock_Name
                    FirstTickerRow = FirstTickerRow - 1
                    If FirstTickerRow = 1 Then Exit Do
            Loop
            
            ' Calculate Open Price of Quarter
            Open_Price = Cells(FirstTickerRow, 3).Value
            Debug.Print "Open Price: " & Open_Price
            
            ' Calculate Close Price of Quarter
            Close_Price = Cells(i, 6).Value
            Debug.Print "Close Price: " & Close_Price
            
            ' Calculate Quarterly Change
            Quarterly_Change = Cells(i, 6).Value - Cells(FirstTickerRow, 3).Value
            Debug.Print "Quarterly Change: " & Quarterly_Change
            
            ' Calculate Percentage Change
            If Cells(FirstTickerRow, 3).Value <> CDbl(0) Then
                Percentage_Change = (Quarterly_Change / Open_Price) * 100
            Else
                Percentage_Change = 0
                Debug.Print "Percentage Change: " & Percentage_Change
            End If
            
            ' Calculate the Stock Total
            Stock_Total = Stock_Total + Cells(i, 7).Value
            Debug.Print "Stock Total: " & Stock_Total
            
            ' Debug output to check values
            Debug.Print "Stock Name: " & Stock_Name
        
            ' Print the Stock Ticker in the Summary Table
            Range("I" & Summary_Table_Row).Value = Stock_Name
            
            ' Print the Quarterly Change in the Summary Table
            Range("J" & Summary_Table_Row).Value = Quarterly_Change
            
            ' Print the Percentage Change in the Summary Table
            Range("K" & Summary_Table_Row).Value = Percentage_Change & "%"
            
            ' Print the Stock Total in the Summary Table
            Range("L" & Summary_Table_Row).Value = Stock_Total
            
             ' Update greatest values
            If Percentage_Change > Max_Percent_Increase Then
                Max_Percent_Increase = Percentage_Change
                Max_Percent_Stock = Stock_Name
            End If
            
            If Percentage_Change < Min_Percent_Decrease Then
                Min_Percent_Decrease = Percentage_Change
                Min_Percent_Stock = Stock_Name
            End If
            
            If Stock_Total > Max_Volume Then
                Max_Volume = Stock_Total
                Max_Volume_Stock = Stock_Name
            End If
            
            ' Move to the next row in the Summary Table
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the totals for the next ticker
            Stock_Total = 0
            Quarterly_Change = 0
            Percentage_Change = 0
            
        Else
            ' Add to the Stock Total for the current ticker
            Stock_Total = Stock_Total + Cells(i, 7).Value
        End If
        
    Next i
    
    'Output Greatest Values we had calculated
    Range("O2").Value = "Greatest % Increase"
    Range("P2").Value = Max_Percent_Stock
    Range("Q2").Value = Max_Percent_Increase
    
    Range("O3").Value = "Greatest % Decrease"
    Range("P3").Value = Min_Percent_Stock
    Range("Q3").Value = Min_Percent_Decrease
    
    Range("O4").Value = "Greatest Total Volume"
    Range("P4").Value = Max_Volume_Stock
    Range("Q4").Value = Max_Volume
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
End Sub

