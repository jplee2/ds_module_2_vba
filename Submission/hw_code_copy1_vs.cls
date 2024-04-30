Sub stock_loop():

    Dim i As Long ' row number
    Dim volume As Double ' contents of column G
    Dim vol_total As Double ' what is going to go in column L
    Dim ticker As String ' what is going to go in column G
    Dim q_change As Double ' going to column J
    Dim p_change As Double ' going to column K
    Dim opening As Double ' used to calculate changes
    Dim closing As Double ' used to calculate changes
    Dim k As Integer ' leaderboard row
    
    Dim RowCount As Long ' to get code to run through all data
    Dim start As Long ' to simplify q_change & p_change code
    Dim j As Integer
    
    ' add functionality
    Dim ws As Worksheet
    For Each ws In Worksheets
        
        ' starting values
        vol_total = 0
        q_change = 0
        p_change = 0
        k = 2
        start = 2
        j = 2
        
        ' add functionality
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        ' begin for loop
        For i = 2 To RowCount:
        
            ' direct our variables to our columns
            volume = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            opening = ws.Cells(i, 3).Value
            closing = ws.Cells(i, 6).Value
            
            ' begin If statement, IF our cell isn't equal to our last cell THEN
            If (ws.Cells(i + 1, 1).Value <> ticker) Then
            
                ' when condition is triggered: calculate our total, q_change, and p_change for our finished ticker array
                vol_total = vol_total + volume
                q_change = ws.Cells(i, 6) - ws.Cells(start, 3)
                p_change = q_change / ws.Cells(start, 3)
                
                ' output our calculations to our leaderboard
                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = q_change
                ws.Cells(k, 11).Value = p_change
                ws.Cells(k, 12).Value = vol_total
                
                ' color our q_change outputs
                Select Case q_change
                        Case Is > 0
                            ws.Range("J" & j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & j).Interior.ColorIndex = 0
                    End Select
                
                ' reset the total and change the leaderboard output row since we just changed ticker types
                vol_total = 0
                k = k + 1
                start = i + 1
                j = j + 1
                
            Else
                ' we just add to the total if the cell's ticker value is the same
                vol_total = vol_total + volume
                
            End If
        Next i
    
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Columns("I:L").AutoFit
    
    ' Add functionality to your script to return the Greatest % Increase, Decrease, and Total
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    ws.Cells(2, 15).Value = ws.Cells(increase_number + 1, 1)
    
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
    ws.Cells(3, 15).Value = ws.Cells(decrease_number + 1, 1)
    
    total_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
    ws.Cells(4, 15).Value = ws.Cells(total_number + 1, 1)
    
    ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(ws.Range("K2:K93001"))
    ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(ws.Range("K2:K93001"))
    ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(ws.Range("L2:L93001"))
    
    Next ws

End Sub
