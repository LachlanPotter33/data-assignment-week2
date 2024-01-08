VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CalculateStockData()
    ' Variables
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Volume As Double
    Dim SummaryRow As Long
    Dim Greatest_Percent_Increase As Variant
    Dim Greatest_Percentage_Decrease As Variant
    Dim Greatest_Total_Volume As Variant
    
    
    For Each ws In ThisWorkbook.Worksheets
    
        ' Adding Columns as Needed
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        SummaryRow = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       '  Ticker Checker
        For i = 2 To LastRow

            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                Ticker = ws.Cells(i, 1).Value

                'Closing Price
                ClosingPrice = ws.Cells(i, 6).Value

                'Yearly Change
                YearlyChange = ClosingPrice - OpeningPrice

                'Percent change
                PercentChange = YearlyChange / OpeningPrice

                'Total stock volume
                TotalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(i - 251, 7), ws.Cells(i, 7)))

                'Summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume

                ' Percentage
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"

                ' Conditional Formatting
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
                End If

                ' Summary table
                SummaryRow = SummaryRow + 1

                ' Opening price
                ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    OpeningPrice = ws.Cells(i, 3).Value
            End If
        Next i
          ' Bonus Question
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
'              'Greatest % Increase
              ws.Cells(2, 17) = "%" & WorksheetFunction.Max(ws.Range("K2:K3001")) * 100
              Greatest_Percent_Increase = WorksheetFunction.Max(ws.Range("K2:K3001"))
              
             P2 = WorksheetFunction.Match(Greatest_Percent_Increase, ws.Range("K2:K3001"), 0)
            ws.Cells(2, 16) = ws.Cells(P2 + 1, 9).Value

'            'Greatest % Decrease
            ws.Cells(3, 17) = "%" & WorksheetFunction.Min(ws.Range("K2:K3001")) * 100
                Greatest_Percent_Decrease = WorksheetFunction.Min(ws.Range("K2:K3001"))

                P3 = WorksheetFunction.Match(Greatest_Percent_Decrease, ws.Range("K2:K3001"), 0)
                ws.Cells(3, 16) = ws.Cells(P3 + 1, 9).Value

'            'Greatest Total Volumne
            ws.Cells(4, 17) = WorksheetFunction.Max(ws.Range("L2:L3001"))
            Greatest_Total_Volume = WorksheetFunction.Max(ws.Range("L2:L3001"))
            
            P4 = WorksheetFunction.Match(Greatest_Total_Volume, ws.Range("L2:l3001"), 0)
                ws.Cells(4, 16) = ws.Cells(P4 + 1, 9).Value

    Next ws
End Sub

