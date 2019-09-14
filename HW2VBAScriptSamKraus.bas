Attribute VB_Name = "Module1"
Sub WallStreet()


    For Each ws In Worksheets
    ws.Activate
    
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        
        Dim Ticker As String
        
        Dim Total_Stock As Double
        Total_Stock = 0
        
        Dim Lastrow As Long
        
        Dim Open_Year_Date As Boolean
        Open_Year_Date = True
        
        Dim Open_Year_Price As Double
        Dim End_Year_Price As Double
        Dim Yearly_Change As Double
        
        Dim Percent_Change As Double
        
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        Lastrow = ws.Cells(Rows.Count, 1).End(x1Up).Row
        
        For i = 2 To Lastrow
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                
                Total_Stock = Total_Stock + Cells(i, 7).Value
                
                Range("I" & Summary_Table_Row).Value = Ticker
                
                Range("L" & Summary_Table_Row).Value = Total_Stock
                
                End_Year_Price = Cells(i, 6).Value
                Yearly_Change = End_Year_Price - Open_Year_Price
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                Percent_Change = (Yearly_Change / Open_Year_Price) * 100
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                
                If Range("J" & Summary_Table_Row).Value > 0 Then
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                Total_Stock = 0
                
                Open_Year_Date = True
                
                
            Else
                Total_Stock = Total_Stock + Cells(i, 7).Value
                
                If Open_Year_Date And Cells(i, 3).Value <> 0 Then
                    Open_Year_Price = Cells(i, 3).Value
                    Open_Year_Date = False
                    
                End If
                
            End If
            
        Next i
        
        Dim n As Integer
        Dim Max_Inc As Long
        Dim Min_Dec As Long
        Max_Inc = 0
        Min_Dec = 0
        Max_Vol = 0
                
        n = ws.Cells(Rows.Count, 10).End(x1Up).Row     
        
        For j = 2 To n
        
            If Range("K" & j).Value > Max_Inc Then
                Max_Inc = Range("K" & j).Value
                Range("P2").Value = Max_Inc
                Range("O2").Value = Range("I" & j).Value
                
            ElseIf Range("K" & j).Value < Min_Dec Then
                Min_Dec = Range("K" & j).Value
                Range("P3").Value = Range("K" & j).Value
                Range("O3").Value = Range("I" & j).Value
                
            End If
            
            If Range("L" & j).Value > Max_Vol Then
                Max_Vol = Range("L" & j).Value
                Range("P4").Value = Max_Vol
                Range("O4").Value = Range("I" & j).Value
                
            End If
            
                Range("P2:P3").NumberFormat = "0.00%"
                Range("P:P").EntireColumn.AutoFit
                
        Next j
        
    Next ws
    
End Sub

