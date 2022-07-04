'====================================================================
'============ CLEAR SUMMARY TABLES FUNCTION
'Function to remove all data and formating from Cells I to Q in all
'worksheets within the Excel file.
'====================================================================
Sub clearSummaryTable()
    For Each ws In Worksheets
        ws.Range("I:Q").Clear
    Next ws
End Sub

'====================================================================
'============= FORMAT STOCK BY TICKER FUNCTION
'Function that will calculate the Yearly Change, Percent Change, and
'Total Volume for each Stock and create a Summary Table in Columns I
'to L of the worksheet. This function will also conditionally color
'code the calculated Yearly Change value: if the value is negative,
'the cell will be colored Red otherwise it will be colored green.
'====================================================================

Sub formatStockTicker()

    Dim tickerName As String      'Variable for holdng the ticker
    
    Dim tickerVol As Double       'Variable for the total Vol per ticker
        tickerVol = 0
        
      
    'Other Variables
    Dim percentChange As Double, yearlyChange As Double
    
    Dim numOfRows As Long     'Variable to find total number of rows in the dataset
        numOfRows = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of the location for each Stock Ticker in the summary table
    Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2       'Summary Table starting in Row 2
    
    '====================================================================
    '==== OPENING AND CLOSING PRICE VARIABLES
    '====================================================================
    Dim openPrice As Double, closingPrice As Double
       
           
    '====================================================================
    '==== HEADER OF THE SUMMARY TABLE
    '====================================================================
     Cells(1, 9) = "Ticker": Cells(1, 10) = "Yearly Change"
     Cells(1, 11) = "Percent Change": Cells(1, 12) = "Total Stock Volume"
     '====================================================================
    
    
    'Loop through all the stocker ticker
           
    For i = 2 To numOfRows
              
        If i = 2 Then
            '====================================================================
            '= SECTION TO GET THE OPENING PRICE FOR FIRST ITEM IN THE TABLE ONLY
            '====================================================================
             'Setting the value of OpenPrice for the first item in the table
             openPrice = Cells(i, 3).Value   
        
        End If
        
        ' If the value in the cells are not the same then...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            '====================================================================
            '==== PART 1: SETTING VARIABLES BEFORE OPENPRICE IS OVERRIDDEN
            '====================================================================
            
            closingPrice = Cells(i, 6).Value
            yearlyChange = closingPrice - openPrice
            percentChange = yearlyChange / openPrice
            
            ' Set the Ticker name
            tickerName = Cells(i, 1).Value
            
            ' Adding the final volume to the Volume Total
            tickerVol = tickerVol + Cells(i, 7).Value
            
            '====================================================================
            '==== PART 2: SETTING VALUES FOR NEXT ITEM IN LIST
            '====================================================================
            
            'Setting the next item's opening Price
            openPrice = Cells(i + 1, 3).Value
            
            '====================================================================
            '=== PRINT TO THE SUMMARY TABLE
            '====================================================================
                
            'Print the Stock Ticker in the Summary Table
             Range("I" & Summary_Table_Row).Value = tickerName
            
            'Print the Ticker Yearly Change to the Summary Table
             Range("J" & Summary_Table_Row).Value = FormatNumber(yearlyChange, 2)
                 If yearlyChange >= 0 Then
                    'Color cell Green if yearly Change is Positive
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    'Color cell red if yearly Change is Negative
                    Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                 End If
             
            'Print the Ticker Percent Change formatted in the Summary Table
             Range("K" & Summary_Table_Row).Value = FormatPercent(percentChange)
            
            'Print the Ticker Volume Total to the Summary Table
             Range("L" & Summary_Table_Row).Value = tickerVol

            '====================================================================
                
             ' Updating Summary Table to the next row by adding one 
             Summary_Table_Row = Summary_Table_Row + 1
    
             ' Reset the Ticker Volume Total
             tickerVol = 0
        
                          
        ' If the cell immediately following a row is the same Ticker...
        Else
          'Add to the Ticker Volume Total
          tickerVol = tickerVol + Cells(i, 7).Value
          
        End If
              
    Next i

End Sub


 '====================================================================
 'BONUS: Functionality to return stock with the "Greatest % Increase",
           '"Greatest % Decrease", and "Greatest Total Volume".
 '====================================================================
Sub greatestValues()
    
    'Since this function is dependent on the table created by the formatStockTicker 
    'function, created a message box to inform the user and give them the opportunity
    'to create the Summary Table first
    
    If Cells(2, 9) = isblank Then
        MsgBox ("Please click the Format Stock by Ticker button first. Thank you.")
    Else
    
        '====================================================================
        '==== PRINT TABLE HEADER AND ROW LABELS FOR COMPARISON TABLE
        '====================================================================
            Cells(1, 16) = "Ticker"
            Cells(1, 17) = "Value"
            Cells(2, 15) = "Greatest % Increase"
            Cells(3, 15) = "Greatest % Decrease"
            Cells(4, 15) = "Greatest Total Volume"
        '====================================================================
    
        Dim i As Integer
        
        
        Dim maxVol As Double       'Variable for the total Vol per ticker
        Dim volumeTickerName As String 'Variable for holdng the ticker with the maximum total volume
            
        
        Dim maxChange As Double         'Variable fro the maximum yearly change
        Dim maxChangeTicker As String   'Variable for holdng the ticker with the maximum yearly change
        
        Dim minChange As Double         'Variable fro the minimum yearly change
        Dim minChangeTicker As String   'Variable for holdng the ticker with the minimum yearly change
        
        
        'Variable to find total number of rows in the Summary Table dataset
        Dim numOfRows As Long           
            numOfRows = Cells(Rows.Count, 9).End(xlUp).Row


        '====================================================================
        'INITALIZING ALL VARIABLES TO THE FIRST ENTRY TO START THE COMPARISON
        '====================================================================    
        maxChange = Cells(2, 11).Value
        minChange = Cells(2, 11).Value
        maxVol = Cells(2, 12).Value
        
        maxChangeTicker = Cells(2, 9).Value
        minChangeTicker = Cells(2, 9).Value
        volumeTickerName = Cells(2, 9).Value
        
        For i = 2 To numOfRows
            'If statement to check for the Max Yearly change
            If Cells(i + 1, 11).Value >= maxChange Then
                maxChange = Cells(i + 1, 11).Value
                maxChangeTicker = Cells(i + 1, 9).Value
            End If
            
            'If statement to check for the Minimum Yearly change
            If Cells(i + 1, 11).Value <= minChange Then
                minChange = Cells(i + 1, 11).Value
                minChangeTicker = Cells(i + 1, 9).Value
            End If
            
            'If statement to check for the Max Total Volume
            If Cells(i + 1, 12).Value >= maxVol Then
                maxVol = Cells(i + 1, 12).Value
                volumeTickerName = Cells(i + 1, 9).Value
            End If
            
        
        Next i
        '====================================================================
        '=== PRINT TO THE SUMMARY TABLE 
        '====================================================================
        Cells(2, 16).Value = maxChangeTicker
        Cells(2, 17).Value = FormatPercent(maxChange)
        
        Cells(3, 16).Value = minChangeTicker
        Cells(3, 17).Value = FormatPercent(minChange)
        
        Cells(4, 16).Value = volumeTickerName
        Cells(4, 17).Value = maxVol
    End If

End Sub


'====================================================================
'=====================  WORKSHEET LOOP FUNCTION =====================
'=COMBINATION OF THE FORMATSTOCKTICKER AND GREATEST VALUE FUNCTIONS.=
'====================================================================
'This function will 
'[1] calculate the Yearly Change, Percent Change, and Total Volume for each Stock 
'[2] create a Summary Table in Columns I to L of the worksheet. 
'[3] conditionally color code the calculated Yearly Change value: if the value 
'is negative, the cell will be colored Red otherwise it will be colored green.
'[4] 'Return stock with the "Greatest % Increase", "Greatest % Decrease",
'and "Greatest Total Volume" for all Worksheet

'====================================================================
Sub WorksheetLoop()

    For Each ws In Worksheets
            Dim tickerName As String      'Variable for holdng the ticker
            
            Dim tickerVol As Double       'Variable for the total Vol per ticker
                tickerVol = 0
                
                        
            'Other Variables
            Dim percentChange As Double, yearlyChange As Double
            
            Dim numOfRows As Long     'Variable to find total number of rows in the dataset
                numOfRows = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Keep track of the location for each Stock Ticker in the summary table
            Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2       'Summary Table starting in Row 2
            
            '====================================================================
            '==== OPENING AND CLOSING PRICE VARIABLES
            '====================================================================
            Dim openPrice As Double, closingPrice As Double
               
                   
            '====================================================================
            '==== HEADER OF THE SUMMARY TABLE
            '====================================================================
              ws.Cells(1, 9) = "Ticker": ws.Cells(1, 10) = "Yearly Change"
              ws.Cells(1, 11) = "Percent Change": ws.Cells(1, 12) = "Total Stock Volume"
            '====================================================================
            
            
            'Loop through all the stocker ticker
                   
            For i = 2 To numOfRows
                      
                If i = 2 Then
                    '====================================================================
                    '= SECTION TO GET THE OPENING PRICE FOR FIRST ITEM IN THE TABLE ONLY
                    '====================================================================
                     'Setting the value of OpenPrice for the first item in the table
                     openPrice = ws.Cells(i, 3).Value    
                
                End If
                
                ' If the value in the cells are not the same then...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                    '====================================================================
                    '==== PART 1: SETTING VARIABLES BEFORE OPENPRICE IS OVERRIDDEN
                    '====================================================================
                    
                    closingPrice = ws.Cells(i, 6).Value
                    yearlyChange = closingPrice - openPrice
                    percentChange = yearlyChange / openPrice
                    
                    ' Set the Ticker name
                    tickerName = ws.Cells(i, 1).Value
                    
                    ' Adding the final volume to the Volume Total
                    tickerVol = tickerVol + ws.Cells(i, 7).Value
                    
                    '====================================================================
                    '==== PART 2: SETTING VALUES FOR NEXT ITEM IN LIST
                    '====================================================================
                    
                    'Setting the next item's opening Price
                    openPrice = ws.Cells(i + 1, 3).Value
                                        
                    '====================================================================
                    '=== PRINT TO THE SUMMARY TABLE
                    '====================================================================
                        
                    'Print the Stock Ticker in the Summary Table
                     ws.Range("I" & Summary_Table_Row).Value = tickerName
                    
                    'Print the Ticker Yearly Change to the Summary Table
                     ws.Range("J" & Summary_Table_Row).Value = FormatNumber(yearlyChange, 2)
                         If yearlyChange >= 0 Then
                            'Color cell Green if yearly Change is Positive
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                        Else
                            'Color cell red if yearly Change is Negative
                            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                         End If
                     
                    'Print the Ticker Percent Change formatted in the Summary Table
                     ws.Range("K" & Summary_Table_Row).Value = FormatPercent(percentChange)
                    
                    'Print the Ticker Volume Total to the Summary Table
                     ws.Range("L" & Summary_Table_Row).Value = tickerVol
        
                    '=====================================================
                          
                     ' Updating Summary Table to the next row by adding one 
                     Summary_Table_Row = Summary_Table_Row + 1
            
                     ' Reset the Ticker Volume Total
                     tickerVol = 0
                
                                  
                ' If the cell immediately following a row is the same Ticker...
                Else        
                  ' Add to the Ticker Volume Total
                  tickerVol = tickerVol + ws.Cells(i, 7).Value
                         
                End If
                               
          Next i
            
            
        '=============================================================
        '==== PRINT TABLE HEADER AND ROW LABELS FOR COMPARISON TABLE
        '=============================================================
            ws.Cells(1, 16) = "Ticker"
            ws.Cells(1, 17) = "Value"
            ws.Cells(2, 15) = "Greatest % Increase"
            ws.Cells(3, 15) = "Greatest % Decrease"
            ws.Cells(4, 15) = "Greatest Total Volume"
        '============================================================
        
        Dim maxVol As Double       'Variable for the total Vol per ticker
        Dim volumeTickerName As String 'Variable for holdng the ticker with the maximum total volume
            
        
        Dim maxChange As Double         'Variable fro the maximum yearly change
        Dim maxChangeTicker As String   'Variable for holdng the ticker with the maximum yearly change
        
        Dim minChange As Double         'Variable fro the minimum yearly change
        Dim minChangeTicker As String   'Variable for holdng the ticker with the minimum yearly change
        
        'Variable to find total number of rows in the Summary Table dataset 
        numOfRows = ws.Cells(Rows.Count, 9).End(xlUp).Row  
            
        '====================================================================
        'INITALIZING ALL VARIABLES TO THE FIRST ENTRY TO START THE COMPARISON
        '==================================================================== 
        maxChange = ws.Cells(2, 11).Value
        minChange = ws.Cells(2, 11).Value
        maxVol = ws.Cells(2, 12).Value
        
        maxChangeTicker = ws.Cells(2, 9).Value
        minChangeTicker = ws.Cells(2, 9).Value
        volumeTickerName = ws.Cells(2, 9).Value
        
        For i = 2 To numOfRows
           'If statement to check for the Max Yearly change
            If ws.Cells(i + 1, 11).Value >= maxChange Then
                maxChange = ws.Cells(i + 1, 11).Value
                maxChangeTicker = ws.Cells(i + 1, 9).Value
                
            End If
            
            'If statement to check for the Minimum Yearly change
            If ws.Cells(i + 1, 11).Value <= minChange Then
                minChange = ws.Cells(i + 1, 11).Value
                minChangeTicker = ws.Cells(i + 1, 9).Value
            End If
            
            'If statement to check for the Max Total Volume
            If ws.Cells(i + 1, 12).Value >= maxVol Then
                maxVol = ws.Cells(i + 1, 12).Value
                volumeTickerName = ws.Cells(i + 1, 9).Value
            End If
            
        
        Next i
         '====================================================================
         '=== PRINT TO THE SUMMARY TABLE 
         '====================================================================
          ws.Cells(2, 16).Value = maxChangeTicker
          ws.Cells(2, 17).Value = FormatPercent(maxChange)
          
          ws.Cells(3, 16).Value = minChangeTicker
          ws.Cells(3, 17).Value = FormatPercent(minChange)
          
          ws.Cells(4, 16).Value = volumeTickerName
          ws.Cells(4, 17).Value = maxVol

    Next ws
End Sub

