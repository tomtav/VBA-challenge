Option Explicit
Sub Stocks()
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
    
        ' Retrieve worksheet's last row and column that is populated
        ' Initialize lastRow and lastCol variables
        Dim lastRow, lastCol As Long
        lastRow = ws.UsedRange.Rows.Count
        lastCol = ws.UsedRange.Columns.Count
        
        ' Initialize summary table current row place holder
        Dim currentSummaryRow As Long
        currentSummaryRow = 2
        
        ' Initalize all variables
        Dim ticker As String
        Dim yearOpen, yearClose, yearlyChange, percentChange, totalVolume As Double
        yearOpen = 0
        yearClose = 0
        yearlyChange = 0
        percentChange = 0
        totalVolume = 0
        
        ' Initialize variables to hold ticker value for each performance metric
        Dim greatestIncStock, greatestDecStock, greatestVolStock As String
        
        ' Initialize variables to store performance metric values and use for comparison
        Dim greatestIncValue, greatestDecValue, greatestVolValue As Double
        greatestIncValue = 0
        greatestDecValue = 0
        greatestVolValue = 0
        
        ' Call sub-routine to generate the summary and performance tables
        SetupSummary ws

        Dim r As Long
        Dim c As Long
        For r = 2 To lastRow
        
            ' Get first instance of stock from sorted column to obtain year open price
            If (ws.Cells(r - 1, 1).Value <> ws.Cells(r, 1).Value) Then
                yearOpen = ws.Cells(r, 3).Value
            End If

            ' Add current row volume to accumulating total volume variable
            totalVolume = totalVolume + ws.Cells(r, 7).Value

            ' Check if last row of current stock reached
            If (ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value) Then
                ticker = ws.Cells(r, 1).Value
                
                yearClose = ws.Cells(r, 6).Value
                yearlyChange = yearClose - yearOpen
                ws.Range("I" & currentSummaryRow).Value = ticker
                ws.Range("J" & currentSummaryRow).Value = yearlyChange

                ' Set conditional CELL color by value of yearlyChange
                If (yearlyChange > 0 ) Then
                    ' if yearlyChange is greater than 0 set CELL to GREEN
                    ws.Range("J" & currentSummaryRow).Interior.ColorIndex = 4
                ElseIf (yearlyChange < 0) Then
                    ' if yearlyChange is less than 0 set CELL color to RED
                    ws.Range("J" & currentSummaryRow).Interior.ColorIndex = 3
                End If

                If (yearOpen = 0) Then
                    percentChange = yearlyChange
                Else
                    percentChange = yearlyChange / yearOpen
                End If
                ws.Range("K" & currentSummaryRow).Value = percentChange
                ws.Range("K" & currentSummaryRow).Style = "Percent"
                
                If (percentChange > 0) Then
                    ' compare percentChange with current stored greatest percent increase value
                    If (percentChange > greatestIncValue) Then
                        greatestIncStock = ticker
                        greatestIncValue = percentChange
                    End If
                ElseIf (percentChange < 0) Then
                    ' compare percentChange with current stored greatest percent decrease value
                    If (percentChange < greatestDecValue) Then
                        greatestDecStock = ticker
                        greatestDecValue = percentChange
                    End If
                End If
                
                ' set total volume cell value in Summary table
                ws.Range("L" & currentSummaryRow).Value = totalVolume
                
                ' compare totalVolume with current stored greatest volume value
                If (totalVolume > greatestVolValue) Then
                    greatestVolStock = ticker
                    greatestVolValue = totalVolume
                End If
                
                ' set number of next row in summary table
                currentSummaryRow = currentSummaryRow + 1
                
                ' reset all counters and aggregators
                yearOpen = 0
                yearClose = 0
                yearlyChange = 0
                percentChange = 0
                totalVolume = 0
                
            End If

        Next r
        
        ' Set performance metrics in designated table
        ws.Range("P2").Value = greatestIncStock
        ws.Range("Q2").Value = greatestIncValue
        ws.Range("Q2").Style = "Percent"
        
        ws.Range("P3").Value = greatestDecStock
        ws.Range("Q3").Value = greatestDecValue
         ws.Range("Q3").Style = "Percent"
        
        ws.Range("P4").Value = greatestVolStock
        ws.Range("Q4").Value = greatestVolValue
        
        ' Set column widths of the summary and performance tables
        ws.Range("I:Q").Columns.AutoFit
        
        'Exit For

    Next

End Sub

Sub SetupSummary(ws As Worksheet)

    ' Ticker analysis headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Font.Bold = True
    
    ' All sheet tickers analysis headers
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greates Total Volume"
    ws.Range("O2:O4").Font.Bold = True
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P1:Q1").Font.Bold = True
    
End Sub

