Attribute VB_Name = "VBA_Challenge2"
Sub Challenge2VBA()

                'Variables
  
    Dim sheet As Worksheet
            
    Dim data_row As Long 'Keep track of Ticker row [I]
    
    Dim row_counter As Long 'Keep track of Ticker row [A]
    
    Dim last_ticker As String
    
    Dim current_ticker As String
    
    Dim first_open As Double
    
    Dim last_close As Double
    
    Dim last_filled_row As Long 'row_counter -1 after while loop is done
    
    Dim percent_change As Double
    
    Dim yearly_change As Double
    
    
            'Apply to all sheets in workbook, independant of the while block
         
    For Each sheet In ActiveWorkbook.Worksheets
            sheet.Activate
            
                    'Column Width for readibility
                    
             Columns("I:L").ColumnWidth = 20
             
                    'Percent Change Number Formatting
                    
             Range("K:K").NumberFormat = "0.00%"
        
                    'Header Names
    
            Cells(1, 9).Value = "Ticker"
            
            Cells(1, 10).Value = "Yearly Change"
            
            Cells(1, 11).Value = "Percent Change"
            
            Cells(1, 12).Value = "Total Stock Volume"
            
            
                    'Values
    
            row_counter = 2
            
            data_row = 2
            
            last_ticker = Cells(row_counter, 1)
        
            first_open = Cells(row_counter, 3).Value
                
            total_volume = 0
        
            Cells(data_row, 9).Value = last_ticker
            
            last_filled_row = row_counter - 1 'Only functions after the while loop and since row_counter is outside the while
            
        
                    'While loop
                        'Will run until the first blank cell in column A, utilizes IsEmpty = False
                        'Will not grab data from the last row of columns [C,F,G] because once it hits the first blank the function is done.
                        'The last row of [C,F,G] is grabbed from seperate lines under Wend
            While IsEmpty(Cells(row_counter, 1)) = False
                current_ticker = Cells(row_counter, 1).Value
            
            'As long as the current ticker is the same, add the values from the <vol> column. Also Add +1 to the row counter.
                If current_ticker = last_ticker Then
                    total_volume = total_volume + Cells(row_counter, 7).Value
                        
                row_counter = row_counter + 1
                
             'If the current ticker is not equal to the last ticker then the last ticker is equal to the row counter
                ElseIf current_ticker <> last_ticker Then
                    last_ticker = Cells(row_counter, 1)
                        Cells(data_row + 1, 9).Value = last_ticker
                    
                    last_close = Cells(row_counter - 1, 6)
                    
                    yearly_change = last_close - first_open
                    
                    If yearly_change < 0 Then
                        Cells(data_row, 10).Interior.ColorIndex = 3
                        
                    Else
                        Cells(data_row, 10).Interior.ColorIndex = 4
                        
                    End If
                    
                    Cells(data_row, 10).Value = yearly_change
                    
                    percent_change = (yearly_change / first_open)
                    
                    Cells(data_row, 11).Value = percent_change
        
                    Cells(data_row, 12).Value = total_volume
                    
                    first_open = Cells(row_counter, 3).Value
                    last_close = 0
                    
                    total_volume = 0
                    
                    data_row = data_row + 1
                End If
                
            Wend
                                  
    'Last row
             last_filled_row = row_counter - 1
             
                last_close = Cells(last_filled_row, 6)
                    
                yearly_change = last_close - first_open
            
                Cells(data_row, 10).Value = yearly_change
            
                percent_change = (yearly_change / first_open)
            
                Cells(data_row, 11).Value = percent_change
        
                Cells(data_row, 12).Value = total_volume
            
    'Color Coding For Yearly Change
            
             If yearly_change < 0 Then
                Cells(data_row, 10).Interior.ColorIndex = 3 'Color index 3 is red.
                
            Else
                Cells(data_row, 10).Interior.ColorIndex = 4 'Color index 4 is green. Black is 1.
                
            End If
            
    'If something goes wrong with the last row
        If IsEmpty(Cells(data_row, 10)) = True Then
            MsgBox "The last rows under Yearly Change -> Total Stock Volume did not populate. For your convinence the sub routine will terminate."
            Exit Sub 'If I don't exit the sub routine the message will show for each sheet. Which is very annoying.
        End If
                    
    Next
     
End Sub


