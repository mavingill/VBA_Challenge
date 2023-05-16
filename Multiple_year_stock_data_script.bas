Attribute VB_Name = "Module1"
Sub multiple_year_stock_data():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        
        Dim tick_counter As Long
        Dim LastRowA As Long
        Dim LastRowI As Long
        
        Dim percent_change As Double
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_volume As Double
        
        Dim i As Long
        Dim j As Long
        
        WorksheetName = ws.Name
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        tick_counter = 2
        j = 2
        
           
        For i = 2 To LastRowA
            
                
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(tick_counter, 9).Value = ws.Cells(i, 1).Value
        ws.Cells(tick_counter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
           
        If ws.Cells(tick_counter, 10).Value < 0 Then
                
        ws.Cells(tick_counter, 10).Interior.ColorIndex = 3
                
        Else
        ws.Cells(tick_counter, 10).Interior.ColorIndex = 4
                
        End If
                    
        If ws.Cells(j, 3).Value <> 0 Then
        percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
        ws.Cells(tick_counter, 11).Value = Format(percent_change, "Percent")
                    
        Else
                    
        ws.Cells(tick_counter, 11).Value = Format(0, "Percent")
                    
        End If
                    
        ws.Cells(tick_counter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
        tick_counter = tick_counter + 1
        j = i + 1
                
        End If
            
        Next i
            
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        greatest_volume = ws.Cells(2, 12).Value
        greatest_increase = ws.Cells(2, 11).Value
        greatest_decrease = ws.Cells(2, 11).Value
        
        For i = 2 To LastRowI
                
        If ws.Cells(i, 11).Value > greatest_increase Then
        greatest_increase = ws.Cells(i, 11).Value
        ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
        Else
                
        greatest_increase = greatest_increase
                
        End If
                
        If ws.Cells(i, 11).Value < greatest_decrease Then
        greatest_decrease = ws.Cells(i, 11).Value
        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
        Else
                
        greatest_decrease = greatest_decrease
                
        End If
        
        If ws.Cells(i, 12).Value > greatest_volume Then
        greatest_volume = ws.Cells(i, 12).Value
        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
        Else
                
        greatest_volume = greatest_volume
                
        End If
                
        ws.Cells(2, 17).Value = Format(greatest_increase, "Percent")
        ws.Cells(3, 17).Value = Format(greatest_decrease, "Percent")
        ws.Cells(4, 17).Value = Format(greatest_volume, "Scientific")
            
        Next i
            
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
        ws.Cells(1, 9).Value = "ticker"
        ws.Cells(1, 10).Value = "yearly_change"
        ws.Cells(1, 11).Value = "percent_change"
        ws.Cells(1, 12).Value = "total_stock_volume"
        ws.Cells(1, 16).Value = "ticker"
        ws.Cells(1, 17).Value = "value"
        ws.Cells(2, 15).Value = "greatest % increase"
        ws.Cells(3, 15).Value = "greatest % decrease"
        ws.Cells(4, 15).Value = "greatest_total_volume"
            
    Next ws
        
End Sub

