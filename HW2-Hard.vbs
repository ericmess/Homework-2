Sub Moderate()
    
    For Each ws In Worksheets
    
        Dim Current_Stock_Ticker As String
        Dim Next_Stock_Ticker As String
        Dim Previous_Stock_Ticker As String
        Dim Stock_Value_Open As Double
        Dim Stock_Value_Closed As Double
        Dim Stock_Value_Change As Double
        Dim Stock_Total_Volume As Double
        Dim Stock_Percentage_Change As Double
        Stock_Total_Volume = 0
        Dim Stock_Value_Total As Double
        Dim Summary_Table_Row As Integer
        Dim Greatest_per_Increase As Double
        Dim Greatest_per_Decrease As Double
        Dim Greatest_Total_Volume As Double
        Dim Current_Percent_Change As Double
        Dim Highest_Percent_Change As Double
        Dim Lowest_Percent_Change As Double
        Dim Current_Total_Stock_Volume As Double
        Dim Highest_Total_Stock_Volume As Double
        Dim Highest_Per_Stock_Ticker As String
        Dim Lowest_Per_Stock_Ticker As String
        Dim Highest_Tot_Stock_Ticker As String
        
        Summary_Table_Row = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow 'Loop to get Total Stock Volume
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Current_Stock_Ticker = ws.Cells(i, 1).Value
                Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = Current_Stock_Ticker
                ws.Range("L" & Summary_Table_Row).Value = Stock_Total_Volume
                Summary_Table_Row = Summary_Table_Row + 1
                Stock_Total_Volume = 0
            Else
                Stock_Total_Volume = Stock_Total_Volume + ws.Cells(i, 7).Value
            End If
        Next i

' End Easy, Start Moderate

        Summary_Table_Row = 2
        For j = 2 To LastRow 'Loop to get Stock_Value_Change
            If ws.Cells(j - 1, 1).Value <> ws.Cells(j, 1).Value Then
                Stock_Value_Open = ws.Cells(j, 3)
            ElseIf ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
                Stock_Value_Closed = ws.Cells(j, 6)
                Stock_Value_Change = Stock_Value_Closed - Stock_Value_Open
                Stock_Percentage_Change = Stock_Value_Change / Stock_Value_Open
                ws.Range("K" & Summary_Table_Row).Value = Stock_Percentage_Change
                ws.Range("J" & Summary_Table_Row).Value = Stock_Value_Change
                Summary_Table_Row = Summary_Table_Row + 1
            End If
        Next j
        
        For k = 2 To LastRow
            ws.Range("K" & k).NumberFormat = "0.00%"
            ws.Range("K:K").EntireColumn.AutoFit
            ws.Range("J" & k).NumberFormat = "0.000000000"
            ws.Range("J:J").EntireColumn.AutoFit
            ws.Range("L:L").EntireColumn.AutoFit
            If ws.Cells(k, 10).Value >= 0 Then
                ws.Range("J" & k).Interior.ColorIndex = 4 'Green
            Else
                ws.Range("J" & k).Interior.ColorIndex = 3 'Red
            End If
        Next k
        
        Current_Percent_Change = 0
        Highest_Percent_Change = 0
        Lowest_Percent_Change = 0
        Current_Total_Stock_Volume = 0
        Highest_Total_Stock_Volume = 0
        Highest_Per_Stock_Ticker = "A"
        Lowest_Per_Stock_Ticker = "A"
        Highest_Tot_Stock_Ticker = "A"
        
        For l = 2 To LastRow
            Current_Percent_Change = ws.Range("K" & l).Value
            Current_Total_Stock_Volume = ws.Range("L" & l).Value
            Current_Stock_Ticker = ws.Range("I" & l).Value
            If Current_Percent_Change > Highest_Percent_Change Then
                Highest_Percent_Change = Current_Percent_Change
                Highest_Per_Stock_Ticker = Current_Stock_Ticker
            End If
            If Current_Percent_Change < Lowest_Percent_Change Then
                Lowest_Percent_Change = Current_Percent_Change
                Lowest_Per_Stock_Ticker = Current_Stock_Ticker
            End If
            If Current_Total_Stock_Volume > Highest_Total_Stock_Volume Then
                Highest_Total_Stock_Volume = Current_Total_Stock_Volume
                Highest_Tot_Stock_Ticker = Current_Stock_Ticker
            End If
        Next l
        
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
        ws.Range("R2").Value = Highest_Percent_Change
        ws.Range("Q2").Value = Highest_Per_Stock_Ticker
        ws.Range("R3").Value = Lowest_Percent_Change
        ws.Range("Q3").Value = Lowest_Per_Stock_Ticker
        ws.Range("R4").Value = Highest_Total_Stock_Volume
        ws.Range("Q4").Value = Highest_Tot_Stock_Ticker
        
        'Greatest_per_Decrease = WorksheetFunction.Min(Range("K:K"))
        'Greatest_per_Increase = WorksheetFunction.Max(Range("K:K"))
        'Greatest_Total_Volume = WorksheetFunction.Max(Range("L:L"))
       
        ws.Range("P:P").EntireColumn.AutoFit
        ws.Range("R:R").EntireColumn.AutoFit
        ws.Range("R2:R3").NumberFormat = "0.00%"
            
    Next ws
End Sub
