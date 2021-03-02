Attribute VB_Name = "Module1"
Sub multi_year()

   For Each ws In Worksheets
   
        'declare variables
        Dim Ticker As String
        Dim open_price As Double
        Dim close_price As Double
        Dim high_price As Double
        Dim low_price As Double
        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim i As Long
    
        Dim data_table As Double
        data_table = 2

        'new columns
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "YearlyChange"
        ws.Range("K1") = "PercentChange"
        ws.Range("L1") = "Volume"


            'column starts at row 2 to last row
            For i = 2 To lastrow

                'check if the ticker is the same value and if not then value is ticker adds up
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                    Ticker = ws.Cells(i, 1).Value
                    
                'values pulled are ticker
                ws.Range("I" & data_table).Value = Ticker

                'closing stock price is value in column 6
                close_price = ws.Cells(i, 6)
        
    
                    'if theres no price, values equal to zero
                    If open_price = 0 Then
                        YearlyChange = 0
                        PercentChange = 0
                    'or else the yearly change = to close-open
                    'percent change = (close/open)/open
                    Else:
                        YearlyChange = close_price - open_price
                        PercentChange = (close_price - open_price) / open_price
                    End If
                
                'format new cells
                ws.Range("L" & data_table).Value = volume
                'volume starts at zero
                volume = 0
                ws.Range("J" & data_table).Value = YearlyChange
                ws.Range("K" & data_table).Value = PercentChange
                ws.Range("K" & data_table).Style = "Percent"
                ws.Range("K" & data_table).NumberFormat = "0.00%"

                'add 1 to data table
                data_table = data_table + 1
                
                'else if the cells dont equal then open price are values in column 3
                ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                open_price = ws.Cells(i, 3)
    
                'else volume is added up in each ticker
                Else: volume = volume + ws.Cells(i, 7).Value

                End If

            Next i
        
                For i = 2 To lastrow
                    'if yearly change is greater than 0 highlight green
                    If ws.Range("J" & i).Value > 0 Then
                    ws.Range("J" & i).Interior.ColorIndex = 4
                    'else make cell red
                    Else:
                        ws.Range("J" & i).Interior.ColorIndex = 3
                    End If

            Next i
    
    Next ws

End Sub
