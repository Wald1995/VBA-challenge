Attribute VB_Name = "Module1"
Sub VBA_Challenge():
    
    'Place the loop first for it can be applied to all the Worksheets
    For Each ws In Worksheets

        Dim WorksheetName As String
        Dim Ticker_Name As String
        Dim Percent_Calculation As Double
        Dim Increase_Calculation As Double
        Dim Decrease_Calculation As Double
        Dim Volume_Total As Double
        
        'Attribute
        WorksheetName = ws.Name
        
        'Create the headers of the results
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'ticker counter
        Ticker_Name = 2
        
        'Start row (beacause in row 1 are the titles)
        j = 2
        
        'find the last row in column A
        LastRow1 = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'loop all the rows
            For i = 2 To LastRow1
            
                'review the ticker name change
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'write the ticker in the resume table
                ws.Cells(Ticker_Name, 9).Value = ws.Cells(i, 1).Value
                
                'calculate and write yearly change in the resume table
                ws.Cells(Ticker_Name, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'conditional for background color
                    If ws.Cells(Ticker_Name, 10).Value < 0 Then
                    
                    'if < 0 put red
                    ws.Cells(Ticker_Name, 10).Interior.ColorIndex = 3
                    
                    Else
                    
                    'if not put green
                    ws.Cells(Ticker_Name, 10).Interior.ColorIndex = 4
                    
                    End If
                    
                   'calculate the percent change in resume table
                    If ws.Cells(j, 3).Value <> 0 Then
                    Percent_Calculation = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                
                    'change format to percent
                    ws.Cells(Ticker_Name, 11).Value = Format(Percent_Calculation, "Percent")
                
                    Else
                
                    ws.Cells(Ticker_Name, 11).Value = Format(0, "Percent")
                    
                    End If
                
                'calculate the total volume in resume table
                ws.Cells(Ticker_Name, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
                'Increase Ticker Name by 1
                Ticker_Name = Ticker_Name + 1
            
                'set new start row of the thicker block
                j = i + 1
            
                End If
                  
            Next i
        
        'find the last row in resume table column I
        LastRow2 = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'variables for sumary
        Increase_Calculation = ws.Cells(2, 11).Value
        Decrease_Calculation = ws.Cells(2, 11).Value
        Volume_Total = ws.Cells(2, 12).Value
        
            'loop for the summary
            For i = 2 To LastRow2
                            
                'for the Iecrease claculation review if the next value is larger. If yes, take the new value and populate ws.cells
                If ws.Cells(i, 11).Value > Increase_Calculation Then
                Increase_Calculation = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Increase_Calculation = Increase_Calculation
                
                End If
                
                'for the Decrease calculation review if the next value is larger. If yes, take the new value and populate ws.cells
                If ws.Cells(i, 11).Value < Decrease_Calculation Then
                Decrease_Calculation = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                Decrease_Calculation = Decrease_Calculation
                
                End If
                
                'review if next value of the total volume is larger. If yes, take the new volume value and populate ws.cells
                If ws.Cells(i, 12).Value > Volume_Total Then
                Volume_Total = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                                
                Else
                
                Volume_Total = Volume_Total
                
                End If

            'Write a summary results in ws.cells
            ws.Cells(2, 17).Value = Format(Increase_Calculation, "Percent")
            ws.Cells(3, 17).Value = Format(Decrease_Calculation, "Percent")
            ws.Cells(4, 17).Value = Format(Volume_Total, "Scientific")
            
            Next i
        
        'adjust column width automatically
        Worksheets(WorksheetName).Columns("A:Z").AutoFit
        
    Next ws
    
End Sub

