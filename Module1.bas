Attribute VB_Name = "Module1"
Sub Stocks()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker_Name As String
    Dim Opening_Value As Double
    Dim Total_Volume As LongLong
    Dim j As Long
    Dim I As Long
    Dim Row As Long
    Dim Greatest As Double
    Dim Lowest As Double
    Dim Greatest_Vol As LongLong

    For Each ws In ThisWorkbook.Worksheets
        'Define Last Row.
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Create Columns.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Lowest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
           
        'Define Variables.
        Ticker_Name = ws.Cells(2, 1).Value
        Opening_Value = ws.Cells(2, 3).Value
        Total_Volume = 0
        
        'Create another index for second set of columns.
        j = 2
        
        'For loop.
        For I = 2 To LastRow
            'If the Ticker in row I does not match Ticker_Name.
            If Ticker_Name <> ws.Cells(I, 1).Value Then
                ws.Cells(j, 9).Value = Ticker_Name
                ws.Cells(j, 10).Value = ws.Cells(I - 1, 6).Value - Opening_Value
                
                'If loop that colors Quarters Change column cell as green, red, or blank.
                If ws.Cells(j, 10).Value > 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(j, 10).Value < 0 Then
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                End If
                
                If Opening_Value <> 0 Then
                    ws.Cells(j, 11).Value = ws.Cells(j, 10).Value / Opening_Value
                Else
                    ws.Cells(j, 11).Value = 0
                End If
                ws.Cells(j, 12).Value = Total_Volume
                
                'Reset variables.
                Ticker_Name = ws.Cells(I, 1).Value
                Opening_Value = ws.Cells(I, 3).Value
                Total_Volume = 0
                
                'Increase index.
                j = j + 1
            
            'If the Ticker in row I does match Ticker_Name.
            Else
                Total_Volume = Total_Volume + ws.Cells(I, 7).Value
            End If
        Next I
        
        'Redefine for Greatest Percentage.
        Ticker_Name = ws.Cells(2, 9).Value
        Greatest = ws.Cells(2, 11).Value
        
        'For loop.
        For Row = 3 To j - 1
            If ws.Cells(Row, 11).Value > Greatest Then
                Greatest = ws.Cells(Row, 11).Value
                Ticker_Name = ws.Cells(Row, 9).Value
            End If
        Next Row
        ws.Cells(2, 16).Value = Ticker_Name
        ws.Cells(2, 17).Value = Greatest
        
        'Redefine for Lowest Percentage.
        Ticker_Name = ws.Cells(2, 9).Value
        Lowest = ws.Cells(2, 11).Value
        
        'For loop.
        For Row = 3 To j - 1
            If ws.Cells(Row, 11).Value < Lowest Then
                Lowest = ws.Cells(Row, 11).Value
                Ticker_Name = ws.Cells(Row, 9).Value
            End If
        Next Row
        ws.Cells(3, 16).Value = Ticker_Name
        ws.Cells(3, 17).Value = Lowest
        
        'Redefine for Greatest Volume.
        Ticker_Name = ws.Cells(2, 9).Value
        Greatest_Vol = ws.Cells(2, 12).Value
        
        'For loop.
        For Row = 3 To j - 1
            If ws.Cells(Row, 12).Value > Greatest_Vol Then
                Greatest_Vol = ws.Cells(Row, 12).Value
                Ticker_Name = ws.Cells(Row, 9).Value
            End If
        Next Row
        ws.Cells(4, 16).Value = Ticker_Name
        ws.Cells(4, 17).Value = Greatest_Vol
        
        ' Format percentage columns
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    Next ws
    
    MsgBox "Analysis complete for all worksheets!", vbInformation
End Sub
