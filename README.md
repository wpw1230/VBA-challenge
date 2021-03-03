# VBA-challenge
Sub VBA_WallStreet()

    ' Create variable for worksheets
    
    Dim ws As Worksheet
    
    ' Loop through all sheets
    For Each ws In Worksheets
    
        ' Add titles to columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Stock Volume"
    
        ' --Print Ticker and Total Stock Volume--
        
        ' Set variable to hold Ticker Name
        Dim Ticker_Name As String
        
        ' Set variable to hold Total Volume per Ticker
        Dim Ticker_Total As Variant
        Ticker_Total = 0
        
        ' Keep track of location for each stock ticker in summary table
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
        
        ' Loop through all trading dates
        For i = 2 To 800000
            
            ' Check if still within same ticker, if not...
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                ' Set ticker name
                Ticker_Name = Cells(i, 1).Value
            
                ' Add to ticker total
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
                ' Print Ticker Name in Summary Table
                Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                ' Print Ticker Total Volume to Summary Table
                Range("J" & Summary_Table_Row).Value = Ticker_Total
                
                ' Add one to summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset Ticker Total
                Ticker_Total = 0
                
            ' If the cell immediately following a row is the same ticker...
            Else
            
                ' Add to Total Volume
                Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
            End If
            
        Next i
            
    Next ws
    
End Sub
