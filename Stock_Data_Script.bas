Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Percent_Change As Double
    Dim Stock_Volume As Long
    Dim Analysis_Table_Row As Long
    Dim LastRow As Long
    Dim J_LastRow As Long
    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double

    
       
    Total_Stock_Volume = 0
    
    
    
    For Each ws In Worksheets
    
        'Analysis Table Columns Name
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Ticker Symbol & Total Stock Volume
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Analysis_Table_Row = 2
        
        Opening_Price = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                Ticker = ws.Cells(i, 1).Value
                
                Closing_Price = ws.Cells(i, 6).Value
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'Print the Ticker in the Analysis Table
                ws.Range("I" & Analysis_Table_Row).Value = Ticker
                
                'Print Yearly Change in the Analysis Table
                Yearly_Change = Closing_Price - Opening_Price
                ws.Range("J" & Analysis_Table_Row).Value = Yearly_Change
                
                'Print Percent Change in the Analysis Table
                If (Yearly_Change <> 0) And (Opening_Price <> 0) Then
                    ws.Range("K" & Analysis_Table_Row).Value = Yearly_Change / Opening_Price
                
                Else
                    ws.Range("K" & Analysis_Table_Row).Value = 0
                    
                End If
                
                'Print the Total Stock Volume to the Analysis Table
                ws.Range("L" & Analysis_Table_Row).Value = Total_Stock_Volume
                              
                Analysis_Table_Row = Analysis_Table_Row + 1
            
                Total_Stock_Volume = 0
                
                'Save Next Ticker Opening Price
                Opening_Price = ws.Cells(i + 1, 3).Value
                
            
            Else
            
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
            End If
                                
        Next i
        
        Analysis_Table_Row = 2
        
        
        'Conditional Formatting - Yearly Change
        
        J_LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        For i = 2 To J_LastRow
        
            If ws.Cells(i, 10).Value < 0 Then
                
                ws.Cells(i, 10).Interior.ColorIndex = 3
                
            Else
            
                ws.Cells(i, 10).Interior.ColorIndex = 4
                
            End If
                
        Next i
        
        ws.Columns("J").NumberFormat = "0.00"
        
        'Conditional Formatting - Percent Change
        
        ws.Columns("K").NumberFormat = "0.00%"

        
        'Challenges - Greatest % increase, Greatest % decrease, Greatest total volume
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
     
        
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & J_LastRow))
        
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & J_LastRow))
        
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & J_LastRow))
        
        For i = 2 To J_LastRow
        
            If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
            
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
            
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
            ElseIf ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
            
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
            End If
        
        Next i
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
   
    Next ws
    

End Sub



