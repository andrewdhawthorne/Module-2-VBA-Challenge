Sub Multi_Year_Stock_Testing()

Dim ws As Worksheet

    For Each ws In Worksheets

        'Declaration of Variables
        Dim Ticker As String
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Total_Volume As Double
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Increase_Percent_Ticker As String
        Dim Decrease_Percent_Ticker As String
        Dim Greatest_Volume_Ticker As String
        Dim Greatest_Percent_Increase As Double
        Dim Greatest_Percent_Decrease As Double
        Dim Greatest_Total_Volume As Double
            
        'Assign value to Variables
        Total_Volume = 0
        Opening_Price = 0
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
           
        'Create Summary Table column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Pull Opening Price
        Dim Opening_Price_Row As Integer
        Opening_Price_Row = 2
        
        'Display Summary Table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
            
        'Loop through all Tickers
        For i = 2 To LastRow
        
            'Pull Opening Value
            If Opening_Price = 0 Then
                
                Opening_Price = ws.Cells(i, 3).Value
            
            End If
            
            'Check if data pertains to same Ticker, and capture next set if not
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set Ticker
                Ticker = ws.Cells(i, 1).Value
            
                'Calculate Yearly Change
                Closing_Price = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                
                'Calculate Percent Change
                Percent_Change = Yearly_Change / Opening_Price
                
                'Add to total volume
                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
            
                'Add results to Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("L" & Summary_Table_Row).Value = Total_Volume
                
                'Add One to the Summary Table Row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset Total Volume
                Total_Volume = 0
                Opening_Price = 0
                
            Else
                
                'Calculate Total Stock Volume
                 Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                
            End If
                 
            
        Next i
            
            'Add conditional formatting to Percent Change
            For i = 2 To LastRow
            
                If ws.Cells(i, 10).Value < 0 Then
                    
                        ws.Cells(i, 10).Interior.ColorIndex = 3
                
                    Else
                
                        ws.Cells(i, 10).Interior.ColorIndex = 4
                
                End If
                
        Next i
        
            'Create "Greatest" Summary Table column headers
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest Percent Increase"
            ws.Cells(3, 15).Value = "Greatest Percent Decrease"
            ws.Cells(4, 15).Value = "Greatest Stock Volume"
            
            'Assign value to Variables
            Greatest_Percent_Increase = 0
            Greatest_Percent_Decrease = 0
            Greatest_Total_Volume = 0
                        
            'Loop through summary
            For i = 2 To LastRow
            
                If ws.Cells(i, 11).Value > Greatest_Percent_Increase Then
                    
                    Greatest_Percent_Increase = ws.Cells(i, 11).Value
                    
                    Increase_Percent_Ticker = ws.Cells(i, 9).Value
            
                End If
        
                If ws.Cells(i, 11).Value < Greatest_Percent_Decrease Then
            
                    Greatest_Percent_Decrease = ws.Cells(i, 11).Value
                    
                    Decrease_Percent_Ticker = ws.Cells(i, 9).Value
            
                End If
                    
                If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
                    
                    Greatest_Total_Volume = ws.Cells(i, 12).Value
            
                    Greatest_Volume_Ticker = ws.Cells(i, 9).Value
            
                End If
            
        Next i
            'Add results to summary table
            ws.Range("P2").Value = Increase_Percent_Ticker
            ws.Range("P3").Value = Decrease_Percent_Ticker
            ws.Range("P4").Value = Greatest_Volume_Ticker
            ws.Range("Q2").Value = Greatest_Percent_Increase
            ws.Range("Q3").Value = Greatest_Percent_Decrease
            ws.Range("Q4").Value = Greatest_Total_Volume
            
            'Add percentage formatting to summary table
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
                            
    Next ws
    
End Sub

