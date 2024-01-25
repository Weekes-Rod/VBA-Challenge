# VBA-Challenge
use VBA scripting to analyze generated stock market data.




Sub YearlyChanges()

    Dim ws As Worksheet
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim Percentage As Double
    Dim TotalStock As LongLong
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim SummaryTableRow As Integer
    Dim LastRow As Long
    
    ' Variables for tracking greatest % increase, decrease, and total volume
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim GreatestVolume As LongLong
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim GreatestVolumeTicker As String
    
    ' Initialize greatest values to initial values
    GreatestIncrease = 0
    GreatestDecrease = 0
    GreatestVolume = 0
    GreatestIncreaseTicker = ""
    GreatestDecreaseTicker = ""
    GreatestVolumeTicker = ""

    ' Loop through each worksheet
    For Each ws In Worksheets
        ' Initialize summary table row
        SummaryTableRow = 2
        
        ' Find the last row of data in column A
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Reset year open
        YearOpen = ws.Cells(2, 3).Value
        
        ' Loop through each row of data
        For i = 2 To LastRow
            ' Check if ticker symbol has changed
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                ' Set variables for the current ticker symbol
                Ticker = ws.Cells(i, 1).Value
                TotalStock = TotalStock + ws.Cells(i, 7).Value
                YearClose = ws.Cells(i, 6).Value
                YearlyChange = YearClose - YearOpen
                
                ' Avoid division by zero error and calculate percentage change
                If YearOpen <> 0 Then
                    Percentage = YearlyChange / YearOpen
                Else
                    Percentage = 0
                End If
                
                ' Add data to summary table
                ws.Cells(SummaryTableRow, 10).Value = Ticker
                ws.Cells(SummaryTableRow, 11).Value = YearlyChange
                ws.Cells(SummaryTableRow, 12).Value = Format(Percentage, "0.00%")
                ws.Cells(SummaryTableRow, 13).Value = TotalStock
                
                ' Update summary table row and reset variables for next ticker symbol
                SummaryTableRow = SummaryTableRow + 1
                YearOpen = ws.Cells(i + 1, 3).Value
                TotalStock = 0
                
                ' Check for greatest % increase, decrease, and total volume
                If Percentage > GreatestIncrease Then
                    GreatestIncrease = Percentage
                    GreatestIncreaseTicker = Ticker
                ElseIf Percentage < GreatestDecrease Then
                    GreatestDecrease = Percentage
                    GreatestDecreaseTicker = Ticker
                End If
                
                If TotalStock > GreatestVolume Then
                    GreatestVolume = TotalStock
                    GreatestVolumeTicker = Ticker
                End If
                
            Else
                ' If ticker symbol is the same, continue adding to TotalStock
                TotalStock = TotalStock + ws.Cells(i, 7).Value
            End If
        Next i
    Next ws
    
    ' Output greatest % increase, decrease, and total volume
    MsgBox "Greatest % Increase: " & Format(GreatestIncrease, "0.00%") & " for Ticker: " & GreatestIncreaseTicker & vbCrLf & _
           "Greatest % Decrease: " & Format(GreatestDecrease, "0.00%") & " for Ticker: " & GreatestDecreaseTicker & vbCrLf & _
           "Greatest Total Volume: " & GreatestVolume & " for Ticker: " & GreatestVolumeTicker

End Sub
