Attribute VB_Name = "Module1"
' ================================================================================================================
'
' Workflow for routine part of homework. Have 3 worksheets containing stock data by year. Need to scroll through
' each worksheet and capture critical information (yearly change in dollars, percent change, total volume of share)
' by ticker symbol. Present this information in a table.
'
' ================================================================================================================

Sub stock_data():

' Code for looping through different worksheets
    For Each ws In Worksheets
            
    ' Initial variable for ticker symbol
        Dim ticker_symbol As String
        
    ' Initial variable for ticker volume
        Dim ticker_volume As Double
        
    ' Initial variable for yearly open ticker value
        Dim yearly_open_ticker_value As Double
        Dim yearly_open_ticker_value_captured As Boolean
            
    ' Initial variable for year end close ticker value
        Dim year_end_close_value As Double
        
    ' Percent Change by Stock
        Dim Percent_Change As Double
        
    ' Location of each stock ticker summary
        Dim Stock_Summary_Table_Row As Integer
        Stock_Summary_Table_Row = 2
        
    ' Variable for calculations
        Dim yearly_change As Double
        Dim yearly_change_percent As Double
        
    ' Determine the Last Row
        Dim Last_Row As Long
        Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    ' Loop through all yearly stock transactions
        
        For i = 2 To Last_Row
        
            ' Set to capture 1st day of trading by year, independent of year starting date.
            ' Plus sets it up to ensure no future daily opening dates are captured

           If yearly_open_ticker_value = False Then
                        
                yearly_open_ticker_value = ws.Cells(i, 3).Value
                yearly_open_ticker_value_captured = True
                                                          
            End If
                                                                
            ' Check to see stock symbol name, year end closing price, and stock trading volume.
            ' Includes condition to eventually capture next ticker symbol year opening price.
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                yearly_open_ticker_value_captured = False
                    
                ticker_symbol = ws.Cells(i, 1).Value
                    
                year_end_close_value = ws.Cells(i, 6).Value
                                                   
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                
            ' Calculation for Yearly Change ($$) amount and yearly change (%%)
            ' In the case of yearly change (%%) need to account for possible scenario were yearly_open_ticker_value equals 0
            ' resulting in a calculation of division by 0. Resolved this with a if-loop.
                                        
                yearly_change = year_end_close_value - yearly_open_ticker_value
                
                If yearly_open_ticker_value > 0 Then
                
                    yearly_change_percent = (year_end_close_value - yearly_open_ticker_value) / yearly_open_ticker_value
                
                Else: yearly_open_ticker_value = 0
                    
                    yearly_change_percent = 0
                    
                End If
                                                                     
            ' Print Summary Table Headers And Color Format
                ws.Range("K1").Value = "Ticker Symbol"
                ws.Range("L1").Value = "Yearly Change"
                ws.Range("M1").Value = "Percent Change"
                ws.Range("N1").Value = "Total Stock Volume"
                ws.Range("K1:N1").Interior.ColorIndex = 15
                ws.Range("K1:N1").Font.Bold = True
                 
            ' Print Stock information in Summary Table. For yearly change represent in currency format. For percent change use %%.
                ws.Range("K" & Stock_Summary_Table_Row).Value = ticker_symbol
                ws.Range("K" & Stock_Summary_Table_Row).HorizontalAlignment = xlCenter
                ws.Range("L" & Stock_Summary_Table_Row).Value = yearly_change
                ws.Range("L" & Stock_Summary_Table_Row).NumberFormat = "$#,##0.00_);($#,##0.00)"
                ws.Range("L" & Stock_Summary_Table_Row).HorizontalAlignment = xlCenter
                ws.Range("M" & Stock_Summary_Table_Row).Value = yearly_change_percent
                ws.Range("M" & Stock_Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("M" & Stock_Summary_Table_Row).HorizontalAlignment = xlCenter
                ws.Range("N" & Stock_Summary_Table_Row).Value = ticker_volume
                ws.Range("N" & Stock_Summary_Table_Row).HorizontalAlignment = xlCenter
        
            ' Adjust column width
                ws.Columns("K:N").AutoFit
                
            ' Loop to highlight / format interior cell color for yearly change. "Green" for positive gains, "Red" for loses.
                If (year_end_close_value - yearly_open_ticker_value) > 0 Then
                    ws.Range("L" & Stock_Summary_Table_Row).Interior.ColorIndex = 4
                
                Else
                    ws.Range("L" & Stock_Summary_Table_Row).Interior.ColorIndex = 3

                End If
                
                Stock_Summary_Table_Row = Stock_Summary_Table_Row + 1
            
            ' Reset the ticker volume
                ticker_volume = 0
                yearly_open_ticker_value = 0
                
            ' If the cell immediately following a row is the same ticker. Add to the ticker volume total
            Else
            
                ticker_volume = ticker_volume + ws.Cells(i, 7).Value
                               
            End If
        
        Next i
 
' ================================================================================================================
'
' Workflow for "Challenge" Part of the Homework. Need to look up the ticker symbol with the
' Greatest_Percent_Increase, Greatest_Percent_Decrease, and Greatest_Total_Volume. And create a
' table that captures this information
'
' ================================================================================================================
 
    ' Initial variable for ticker symbol
        Dim ticker_symbol2 As String
        Dim ticker_symbol_percent_increase As String
        Dim ticker_symbol_percent_decrease As String
        Dim ticker_symbol_total_volume As String
        
    ' Initial variable for performance factors on yearly basis
        Dim Greatest_Percent_Increase As Double
        Dim Greatest_Percent_Decrease As Double
        Dim Greatest_Total_Volume As Double
       
    ' Print Summary Table Headers and Format Cells (color and font position)
        ws.Range("R1").Value = "Ticker Symbol"
        ws.Range("R1").HorizontalAlignment = xlCenter
        ws.Range("S1").Value = "Value"
        ws.Range("S1").HorizontalAlignment = xlCenter
        ws.Range("R1:S1").Interior.ColorIndex = 15
        ws.Range("R1:S1").Font.Bold = True
        ws.Range("Q2").Value = "Greatest % Increase"
        ws.Range("Q3").Value = "Greatest % Decrease"
        ws.Range("Q4").Value = "Greatest Total Volume"
        ws.Range("Q2:Q4").Interior.ColorIndex = 15
        ws.Range("Q2:Q4").Font.Bold = True
    
    ' Adjust column width to fit font
        ws.Columns("Q:S").AutoFit
    
    ' Determine the Last Row of Ticker Symbols
        Dim Last_Row2 As Long
        Last_Row2 = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        Greatest_Percent_Increase = WorksheetFunction.Max(ws.Range("M2:M" & Last_Row2))
        Greatest_Percent_Decrease = WorksheetFunction.Min(ws.Range("M2:M" & Last_Row2))
        Greatest_Total_Volume = WorksheetFunction.Max(ws.Range("N2:N" & Last_Row2))

    ' Loop for stock ticker look up for Greatest_Percent_Increase, Greatest_Percent_Decrease, and Greatest_Total_Volume
        
        For j = 2 To Last_Row2
            
            If ws.Cells(j + 1, 13).Value = Greatest_Percent_Increase Then
        
                    ticker_symbol_percent_increase = ws.Cells(j + 1, 11).Value
            
            End If
            
            If ws.Cells(j + 1, 13).Value = Greatest_Percent_Decrease Then
        
                    ticker_symbol_percent_decrease = ws.Cells(j + 1, 11).Value
            
            End If
            
            If ws.Cells(j + 1, 14).Value = Greatest_Total_Volume Then
        
                    ticker_symbol_total_volume = ws.Cells(j + 1, 11).Value
            
            End If
        
        Next j
        
    ' Print Stock information in Summary Table. For yearly change represent in currency format. For percent change use %%.
                ws.Range("S2").Value = Greatest_Percent_Increase
                ws.Range("S2").NumberFormat = "0.00%"
                ws.Range("S2").HorizontalAlignment = xlCenter
                ws.Range("S3").Value = Greatest_Percent_Decrease
                ws.Range("S3").NumberFormat = "0.00%"
                ws.Range("S3").HorizontalAlignment = xlCenter
                ws.Range("S4").Value = Greatest_Total_Volume
                ws.Range("S4").HorizontalAlignment = xlCenter
                ws.Range("R2").Value = ticker_symbol_percent_increase
                ws.Range("R2").HorizontalAlignment = xlCenter
                ws.Range("R3").Value = ticker_symbol_percent_decrease
                ws.Range("R3").HorizontalAlignment = xlCenter
                ws.Range("R4").Value = ticker_symbol_total_volume
                ws.Range("R4").HorizontalAlignment = xlCenter
    
    'Moves to next worksheet in workbook
    Next ws
    
End Sub



