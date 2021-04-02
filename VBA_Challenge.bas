Attribute VB_Name = "VBA_Challenge"
Sub consolidate_data()
    'Loop through all sheets
        For Each ws In Worksheets
          
            Dim current_ticker As String
            Dim i As Long
            Dim volume_total As Double
            Dim yearly_opening As Double
            Dim yearly_closing As Double
            Dim print_counter As Long
            
            
            'Formate new cells to track data
            ws.Range("I1").Value = "Ticker"
            
            ws.Range("J1").Value = "Yearly Change"
            ws.Columns(10).NumberFormat = "#,##0.00"
            
            ws.Range("K1").Value = "Percent Change"
            ws.Columns(11).NumberFormat = "0.00%"
            
            ws.Range("L1").Value = "Total Stock Volume"
            ws.Columns(12).NumberFormat = "#,##0"
            
            
            'initiate print_counter to 2 for each new sheet
            print_counter = 2
            
            'initiate i to 2 so that it references A2 in first while coniditional check
            i = 2
            
            
            'while there is data in the sheet
             While ws.Cells(i, 1).Value <> ""
            
                'initiate volume_total to 0
                
                volume_total = 0
                
                'initiate current_ticker
                current_ticker = ws.Cells(i, 1).Value
                
                'populate yearly_opening
                yearly_opening = ws.Cells(i, 3).Value
                'MsgBox (current_ticker & " yearly_opening is " & yearly_opening)
                
                
                'while current ticker = next ticker
                While current_ticker = ws.Cells(i + 1, 1).Value
                    
                    'aggregate volume
                    volume_total = volume_total + ws.Cells(i, 7).Value
                    
                    'increase i to refer to next cell
                    i = i + 1
                
                Wend
                
                    'add last volume to total
                    volume_total = volume_total + ws.Cells(i, 7)
                    
                    yearly_closing = ws.Cells(i, 6).Value
                    
                    
                    'increase i to refer to next cell
                    i = i + 1
                
                    '------------------------------------
                    'print/format data to fill new columns: ticker, yearly change, percent change, and total stock volume
                    '------------------------------------
                    ws.Cells(print_counter, 9).Value = current_ticker
                    ws.Cells(print_counter, 10).Value = yearly_closing - yearly_opening
                    
                    If yearly_closing And yearly_opening <> 0 Then
                        ws.Cells(print_counter, 11).Value = (yearly_closing / yearly_opening) - 1
                    Else
                        ws.Cells(print_counter, 11).Value = 0
                    End If
                    
                    ws.Cells(print_counter, 12).Value = volume_total
                    
                    'format color of yearly change column
                    If ws.Cells(print_counter, 10).Value > 0 Then
                        ws.Cells(print_counter, 10).Interior.Color = vbGreen
                    ElseIf ws.Cells(print_counter, 10).Value < 0 Then
                        ws.Cells(print_counter, 10).Interior.Color = vbRed
                    End If
                    
                    'increase print_counter to refer to next empty row on combined_tickers
                    print_counter = print_counter + 1
            
            Wend
            
            
            
            '----------------------------------------
            'Create new cells showing greating % increase/decrease and greatest total volume
            '----------------------------------------
            
            'print data identifiers
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            'declare useful variables
            Dim greatest_percent As Double
            Dim least_percent As Double
            Dim greatest_volume As Double
            Dim greatest_ticker As String
            Dim least_ticker As String
            Dim volume_ticker As String
            Dim j As Long
            
            'initiate variables
            greatest_percent = 0
            least_percent = 0
            greatest_volume = 0
            j = 2
            
            'while loop to find greatest/least percent values
            While ws.Cells(j, 11).Value <> ""
                
                'if and elseif to find greatest percent increase and decrease
                If ws.Cells(j, 11).Value > greatest_percent Then
                    greatest_percent = ws.Cells(j, 11).Value
                    greatest_ticker = ws.Cells(j, 9).Value
                    
                ElseIf ws.Cells(j, 11).Value < least_percent Then
                    least_percent = ws.Cells(j, 11).Value
                    least_ticker = ws.Cells(j, 9).Value
                End If
                    
            j = j + 1
            Wend
            
            'reset j
            j = 2
            
            'while loop to find greatest_volume:
            '*******i had this same code in the previous loop, attempting to do
            '*******the work all on one loop, but for some reason the greatest_volume if statement was always evaluating
            '*******to true. Creating its own while loop fixed the issue but i still dont know why the follow if statement wasnt
            '*******working in the previous while loop
            While ws.Cells(j, 12).Value <> ""
                If ws.Cells(j, 12).Value > greatest_volume Then
                    greatest_volume = ws.Cells(j, 12).Value
                    volume_ticker = ws.Cells(j, 9).Value
                End If
            j = j + 1
            Wend
            
            'print values greatest percent increase, greatest percent decrease, and greatest total volume
            ws.Range("P2").Value = greatest_ticker
            ws.Range("Q2").Value = greatest_percent
            ws.Range("Q3").Value = least_percent
            ws.Range("P3").Value = least_ticker
            ws.Range("P4").Value = volume_ticker
            ws.Range("q4").Value = greatest_volume
            
           
           
           '----------------------------------------------
            'Final Formatting before going to next worksheet
            '-----------------------------------------------
            'autofit to display data
            ws.Columns("I:Q").AutoFit
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            
            
        'go to next worksheet
        Next ws
                                         
End Sub

