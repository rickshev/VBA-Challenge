Attribute VB_Name = "Module1"
Sub stocks()
    
    'for loop to run through all worksheets
    For Each ws In Worksheets
    
        'headers for all worksheets
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
        'Declare variables for analysis
        Dim lastRow As Long
        Dim total As Double
        Dim summary As Integer
        Dim first As Double
        Dim last As Double
        Dim x As Long
        
    
        'Set variable values
        lastRow = Cells(Rows.Count, "A").End(xlUp).Row
        summary = 2
        total = 0
        x = 2
    
        'For loop to read through column "A" and find non-equal Strings
        For i = 2 To lastRow
        
            'add total ticker volume
            total = total + ws.Cells(i, 7).Value
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'filling summary ticker data
            ws.Cells(summary, 9).Value = ws.Cells(i, 1).Value
            'total ticker volume
            ws.Cells(summary, 12).Value = total
            'reset ticker total
            total = 0
            
            'first = opening stock price per year; last = closing stock price per year
            first = ws.Cells(x, 3).Value
            last = ws.Cells(i, 6).Value
            
            'yearly change
            ws.Cells(summary, 10).Value = first - last
            
            'color coordinate postiive and negative values
            If ws.Cells(summary, 10).Value >= 0 Then
                ws.Cells(summary, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(summary, 10).Interior.ColorIndex = 3
            End If
            
            'percent change
            If first = 0 Then
                ws.Cells(summary, 11).Value = 0
            Else
                ws.Cells(summary, 11).Value = (ws.Cells(summary, 10).Value) / first
            End If
            
            'format percent change to percentage data
            ws.Cells(summary, 11).NumberFormat = "0.00%"
        
            'new summary line for new ticker
            summary = summary + 1
        
            End If
        Next i
        
        'greatest increase, decrease, and volume - reset lastRow value
        lastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
        'for loop for final results
        For i = 2 To lastRow
        
            'greatest % increase
            If ws.Cells(i, 11).Value > ws.Range("Q2").Value Then
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            End If
            
            'greatest % decrease
            If ws.Cells(i, 11).Value < ws.Range("Q3").Value Then
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            End If
            
            'greatest volume
            If ws.Cells(i, 12).Value > ws.Range("Q4").Value Then
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
            
            'format increase / decrease as percentages
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
            
        Next i
        
    Next ws
    
End Sub

