Sub LoopYear()
    Dim ws As Worksheet
    Dim Ticker_Name As String
    ' Set an initial variable for holding the total volume per Ticker value
    Dim Volume_Total As Double
    Volume_Total = 0
    Dim LastRow As Long
    ' Keep track of the location for each Ticker value in the summary table
    Dim Ticker_Row As Long
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
        'initiate TR from zero
        Ticker_Row = 2
        'Enter the column headers for each sheet
        ws.Range("I" & Ticker_Row - 1).Value = "Ticker"
        ws.Range("J" & Ticker_Row - 1).Value = "Total Stock Volume"
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            ' Loop through all Ticker values
            For i = 2 To LastRow
        
                ' Check if we are still within the same Ticker value, if it is not...
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                  ' Set the Ticker name
                  Ticker_Name = ws.Cells(i, 1).Value
            
                  ' Add to the volume Total
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
                  ' Print the Ticker name in the Combined Ticker column
                  ws.Range("I" & Ticker_Row).Value = Ticker_Name
            
                  ' Print the Ticker Amount to the Summary Table
                  ws.Range("J" & Ticker_Row).Value = Volume_Total
            
                  ' Add one to the summary table row
                  Ticker_Row = Ticker_Row + 1
                  
                  ' Reset the Brand Total
                  Volume_Total = 0
            
                ' If the cell immediately following a row is the same brand...
                Else
            
                  ' Add to the volume Total
                  Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            
                End If
        
            Next i
    Next ws

End Sub



