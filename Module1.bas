Attribute VB_Name = "Module1"


Sub Stock_market()

    'declare and set worksheet
    Dim ws As Worksheet
    Dim Ticker As String
    Dim Ticker_count As Long
    Dim Ticker_volume As Variant
    Dim open_price As Double
    Dim close_price As Double
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim i As Long
    
    For Each ws In Worksheets
        ws.Activate
        'create column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly change"
        ws.Range("k1").Value = " Percent change"
        ws.Range("L1").Value = "Total Stock volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        'Define Ticker variable
         
    
        Ticker = ""
    
        
        'Set new variables for prices and percent changes
        
        open_price = 0
       
        close_price = 0
        
        percent_change = 0
      
        yearly_change = 0
        
        'Create variable to hold stock volume
        
        Ticker_volume = 0

        'Define Lastrow of worksheet
        Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' keep track of ticker in the summary table
        Dim summary_table_row As Integer
        summary_table_row = 2
        'loop of current worksheet to Lastrow
        For i = 2 To Lastrow
            ' check to see if are within the same ticker, if not
            If Ticker <> ws.Cells(i, 1).Value Then
                If Ticker <> "" Then
                
                    yearly_change = close_price - open_price
                    percent_change = (yearly_change / open_price)
                    ' print ticker in the summary table
                    ws.Range("I" & summary_table_row).Value = Ticker
                    'Print the Volume of stocks to the Summary Table
                    ws.Range("L" & summary_table_row).Value = Ticker_volume
                    
                    'Print the yearly change in the Summary Table
                    ws.Range("J" & summary_table_row).Value = yearly_change
                    
                    'Print the percent change in the Summary Table
                    
                    ws.Range("K" & summary_table_row).Value = percent_change
                    'Add one to the summary table row
                    summary_table_row = summary_table_row + 1
                    
                End If
    
                'Ticker Name
                Ticker = ws.Cells(i, 1).Value
                
                'add the volume of ticker
                Ticker_volume = ws.Cells(i, 7).Value
                
                'Calculate change in Price
                close_price = ws.Cells(i, 6).Value
                open_price = ws.Cells(i, 3).Value
                
                
            Else
                Ticker_volume = Ticker_volume + ws.Cells(i, "G")
                close_price = ws.Cells(i, 6).Value
            End If
            

        Next i
        yearly_change = close_price - open_price
        percent_change = (yearly_change / open_price)
        ' print ticker in the summary table
        ws.Range("I" & summary_table_row).Value = Ticker
        'Print the Volume of stocks to the Summary Table
        ws.Range("L" & summary_table_row).Value = Ticker_volume
        
        'Print the yearly change in the Summary Table
        ws.Range("J" & summary_table_row).Value = yearly_change
        
        'Print the percent change in the Summary Table
        
        ws.Range("K" & summary_table_row).Value = percent_change
        
        'additional for loop to go through the summary table and '
         For i = 2 To summary_table_row
         
         ' determine Greatest % Increase
          If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
             ws.Range("Q2").Value = ws.Range("K" & i).Value
             ws.Range("P2").Value = ws.Range("I" & i).Value
             
            End If
            
        
         ' determine Greatest % decrease
         
         If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Range("K" & i).Value
            ws.Range("P3").Value = ws.Range("I" & i).Value
        
         End If
         
         
         ' determine Greatest Total Volume
         
           
         If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Range("L" & i).Value
            ws.Range("P4").Value = ws.Range("I" & i).Value
            
            End If
            
            Next i
            
    
            Next ws
    
    



End Sub








