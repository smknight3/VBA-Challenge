VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockTesting():

    'Set an initial variable for the worksheet
    Dim ws As Worksheet
    
    'Set an initial variable for holding the ticker
    Dim Ticker_Name As String

    'Set an initial variable for holding the total volume of each type of stock
    Dim Stock_Volume As Double
    Stock_Volume = 0
    
    'Set an initial variable for the first date
    Dim First_Price As Double
    
    Dim i As Double
    
    'Set an initial variable for the last date
    Dim Last_Date As Double
    
    'Set an initial variable for the closing price
    Dim Last_Price As Double
    
    'Set an initial variable for the yearly change
    Dim Yearly_Change As Double
    
    'Set an initial variable for the percent change
    Dim Percent_Change As Double
    
    'Set initial variables for second summary table
    Dim Increase_Ticker As String
    Dim Decrease_Ticker As String
    Dim Volume_Ticker As String
    
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    

    'Keep track of the location for each kind of stock in the summary table starting on the second row
    Dim Summary_Table_Row As Integer

    'Loop through the multiple worksheets on the file
    For Each ws In Worksheets
    
        'Define location of summary table for each workbook
        Summary_Table_Row = 2

        'Set the first price
        First_Price = ws.Cells(2, 3).Value
        
        'Set the last date
        Last_Date = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Set titles
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Total Stock Volume"
        ws.Range("K1").Value = "Yearly Changes"
        ws.Range("L1").Value = "Percent Changes"
        ws.Range("N2").Value = "Greatest Percent Increase"
        ws.Range("N3").Value = "Greatest Percent Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"
        
        'Loop through
        For i = 2 To Last_Date
        
            'Calculate sum of stock volume for each stock type
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
            'Check if we are still within the stock type
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'Set the Stock name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Place the stock names in the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
               
                'Place the sum of stock volumes in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Stock_Volume
                
                'Find closing price
                Last_Price = ws.Cells(i, 6).Value
            
                'Find yearly change
                Yearly_Change = Last_Price - First_Price
                
                'Place yearly change on summary table
                ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
            
                'Find percent change
                If First_Price = 0 Then
                    Percent_Change = 100 * (Yearly_Change / 1E-07)
                Else:
                    Percent_Change = 100 * (Yearly_Change / First_Price)
                End If
                
                'Place percent change on summary table
                ws.Range("L" & Summary_Table_Row).Value = Percent_Change

                'Conditional to color positive and negative yearly changes
                If Yearly_Change > 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf Yearly_Change < 0 Then
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
                Else
                    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 2
                End If
                
                'Position summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                Stock_Volume = 0
                First_Price = ws.Cells(i + 1, 3)
                
            End If
    
        Next i
        
        'Loop through summary table
        For i = 2 To Summary_Table_Row
        
            'Find greatest percentage increase
            If ws.Cells(i, 11).Value > Greatest_Increase Then
                Greatest_Increase = ws.Cells(i, 11).Value
                Increase_Ticker = ws.Cells(i, 9).Value
            End If
            
            'Find greatest percentage decrease
            If ws.Cells(i, 11).Value < Greatest_Decrease Then
                Greatest_Decrease = ws.Cells(i, 11).Value
                Decrease_Ticker = ws.Cells(i, 9).Value
            End If
            
            'Find greatest volume
            If ws.Cells(i, 12).Value > Greatest_Volume Then
                Greatest_Volume = ws.Cells(i, 12).Value
                Volume_Ticker = ws.Cells(i, 9).Value
            End If
            
        Next i
        
        'Paste to second summary
        ws.Range("O2").Value = Increase_Ticker
        ws.Range("O3").Value = Decrease_Ticker
        ws.Range("O4").Value = Volume_Ticker
        ws.Range("P2").Value = Greatest_Increase
        ws.Range("P3").Value = Greatest_Decrease
        ws.Range("P4").Value = Greatest_Volume
    
    Next ws
    
End Sub

