Attribute VB_Name = "Module2"
Sub vba_challenge()

'------------------------------
    'Loop for each worksheet
'------------------------------
    For Each ws In Worksheets
        

    '-------------------------
        'DEFINING VARIABLES
    '-------------------------
    
        'Use lastrow function to count lastrow in each wksheet
        Dim lastrow As Double
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
                
        'Define Summary Table Row to keep track of location of each ticker in summary table
        'Equals to 2 because we want to start after header row
        Dim SummaryTableRow As Double
            SummaryTableRow = 2
        
        'Define ticker as a String
        Dim Ticker As String
        
        'Define yearly change variables to help with calculations later
        Dim StartYear As Double
        Dim EndYear As Double
        'Define yearly change as Double (to store more data so it doesn't have overflow error)
        Dim YearlyChange As Double
        
        'Define Percent Change as a double (to store more data so it doesn't have overflow error)
        Dim PercentChange As Double
        
        'Define Total Stock Volume as double (to store more data so it doesn't have overflow error)
        Dim TotalStockVol As Double
            'Set TotalStockVol equal to 0 to reset for every worksheet
            TotalStockVol = 0
        
        'Want to use a ticker counter to be able to compare the very first ticker row opening with the very last ticker row closing. This will come in handy later
        Dim TickerCount As Double
            'Set Ticker count to 0 initially so it resets per ticker
            TickerCount = 0
            
    '---------------------------------
        'FORMATTING EACH WORKSHEET
    '---------------------------------
        'Add column headers for summary table
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Volume of Stock"
        
        'Add column headers for bonus table
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest total volume"
            
        'Format column L as percentage for percent change
        ws.Range("L:L").NumberFormat = "0.00%"
        'Format cells Q2 and Q3 as percentage for greatest % increase and decrease
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        'Format Column M as number for total volume of stock
        ws.Range("M:M").NumberFormat = "0.00"
        'Format cell Q4 as number for greatest stock volume
        ws.Range("Q4").NumberFormat = "0.00"
        
        'Conditional formatting of Yearly Change for positive and negative change
        'Delete previous conditional formats to start fresh
        ws.Range("K2:K" & lastrow).FormatConditions.Delete
                
        'If cell value is BLANK then remove ALL CELL FORMATTING
        ws.Range("K2:K" & lastrow).FormatConditions.Add Type:=xlExpression, Formula1:= _
                                       "=LEN(TRIM(K2))=0"
        ws.Range("K2:K" & lastrow).FormatConditions(1).Interior.Pattern = xlNone
            
        'If cell value is LESS THAN 0 then colour it RED
        ws.Range("K2:K" & lastrow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                                       Formula1:="=0.00"
        ws.Range("K2:K" & lastrow).FormatConditions(2).Interior.ColorIndex = 3
            
        'If cell value is MORE THAN OR EQUAL TO 0 then colour it GREEN
        ws.Range("K2:K" & lastrow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, _
                                       Formula1:="=0.00"
        ws.Range("K2:K" & lastrow).FormatConditions(3).Interior.ColorIndex = 4
                

                
    '----------------
        'ROW LOOP
    '----------------
        For r = 2 To lastrow
           
            'If the next ticker cell AFTER the current cell does NOT equal the current cell...
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then

            '---------------------
                'TICKER COLUMN
            '---------------------
                'Set the Ticker to equal the value in row r column 1
                Ticker = ws.Cells(r, 1).Value
                
                'Print the Ticker into the first summary table row (row 2 as we defined)
                ws.Cells(SummaryTableRow, 10).Value = Ticker
                
            '----------------------------
                'YEARLY CHANGE COLUMN
            '----------------------------
                'Define column <open> as StartYear FOR CURRENT ROW MINUS TICKER COUNT (to get initial row where ticker was first identified)
                'and column <close> as EndYear FOR CURRENT ROW
                StartYear = ws.Cells(r - TickerCount, 3).Value
                EndYear = ws.Cells(r, 6).Value
                
                'Define YearlyChange as EndYear minus StartYear
                YearlyChange = EndYear - StartYear
                
                'Print YearlyChange into Column K
                ws.Cells(SummaryTableRow, 11).Value = YearlyChange
                
            '----------------------------
                'PERCENT CHANGE COLUMN
            '----------------------------
                'Calculate percentage of yearly change from opening price at the beginning of year to closing price at end of year
                'To do this we want to find percentage of the yearly change calculated above divided by year start
                ws.Cells(SummaryTableRow, 12).Value = YearlyChange / StartYear
                
            '---------------------------------
                'TOTAL STOCK VOLUME COLUMN
            '---------------------------------
                'Add current stock volume to previous stock volume and print this value
                TotalStockVol = TotalStockVol + ws.Cells(r, 7).Value
                ws.Cells(SummaryTableRow, 13).Value = TotalStockVol
                
            '------------------------
                'SUMMARY TABLE ROW
            '------------------------
                'Redefine the summary table row so it uses the next row down
                SummaryTableRow = SummaryTableRow + 1
                
            '-------------------------------
                'STOCK VOLUME TOTAL RESET
            '-------------------------------
                'Reset total stock volume for the next one
                TotalStockVol = 0
                
            '---------------------
                'TICKER COUNTER
            '---------------------
                'Reset Ticker Count to 0 for the next Ticker
                TickerCount = 0
                
            'If the cell immediately following the current row is the same ticker then...
            
            Else
                'Add current stock volume to previous stock volume
                TotalStockVol = TotalStockVol + ws.Cells(r, 7).Value
                
                'Add one to the ticker count
                TickerCount = TickerCount + 1
                
            End If
            
        Next r
        
    '------------------
        'BONUS TABLE
    '------------------
    
        'Define variables for greatest % increase, % decrease, and greatest stock volume
        Dim PercentIncreaseTicker As String
        Dim PercentIncreaseValue As Double
            PercentIncreaseValue = 0
        
        Dim PercentDecreaseTicker As String
        Dim PercentDecreaseValue As Double
            PercentDecreaseValue = 0
        
        Dim StockVolTicker As String
        Dim StockVolValue As Double
            StockVolValue = 0
        
        'Need to create a new lastrow variable for the summary table rahter than data table
        lastrow_bonus = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'First loop for greatest % increase
        For percentincrease = 2 To lastrow_bonus
            
            'If current cell value is larger than the previous Percentage Increase Value...
            If ws.Cells(percentincrease, 12).Value > PercentIncreaseValue Then
                    
                'Redefine Percent Increase Ticker name and value from summary table to correspond to that row
                PercentIncreaseTicker = ws.Cells(percentincrease, 10).Value
                PercentIncreaseValue = ws.Cells(percentincrease, 12).Value
                    
            End If
        
        Next percentincrease
        

        'Second loop for greatest % decrease
        For percentdecrease = 2 To lastrow_bonus
        
            'If current cell value is LESS than previous percentage DECREASE value...
            If ws.Cells(percentdecrease, 12).Value < PercentDecreaseValue Then
                    
                'Redefine Percent Increase Ticker name and value from summary table to correspond to that row
                PercentDecreaseTicker = ws.Cells(percentdecrease, 10).Value
                PercentDecreaseValue = ws.Cells(percentdecrease, 12).Value
                
            End If
                
        Next percentdecrease
        
        'New loop for greatest stock volume
        For stockvol = 2 To lastrow_bonus
            
            'If current cell value is larger than previous Greatest Stock Volume...
            If ws.Cells(stockvol, 13).Value > StockVolValue Then
                StockVolTicker = ws.Cells(stockvol, 10).Value
                StockVolValue = ws.Cells(stockvol, 13).Value
                
            End If
                         
        Next stockvol
        
    'Populate the bonus table with the newly defined %increase, %decrease, and greatest stock volume values
        ws.Range("P2").Value = PercentIncreaseTicker
        ws.Range("Q2").Value = PercentIncreaseValue
            
        ws.Range("P3").Value = PercentDecreaseTicker
        ws.Range("Q3").Value = PercentDecreaseValue
            
        ws.Range("P4").Value = StockVolTicker
        ws.Range("Q4").Value = StockVolValue
        
        'Autofit all column widths to display data correctly
        ws.Range("J:Q").EntireColumn.AutoFit
    Next ws
        
End Sub


