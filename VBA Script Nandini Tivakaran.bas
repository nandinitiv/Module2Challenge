Attribute VB_Name = "Module3"
Sub stock_analysis():

For Each ws In Worksheets

'Declaring variables
  Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double

    Dim currentTicker As String
    Dim nextTicker As String
    

        j = 0
        total = 0
        change = 0
        start = 2
               
        
'Naming columns
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'find the row number of the last row with data
rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

' Go through the whole data set starting at row 2 until the last row
' If ticker changes then print results

For i = 2 To rowCount
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'Store results in variables
total = total + ws.Cells(i, 7).Value
ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value

'Handle zero total volume to avoid dividing by zero
If total = 0 Then
percentChange = 0

'print the results

ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
ws.Range("J" & 2 + j).Value = change
ws.Range("K" & 2 + j).Value = percentChange
ws.Range("L" & 2 + j).Value = total

j = j + 1
start = i + 1
total = 0
change = 0

Else
percentChange = change / total

 ' Find First non zero starting value
 
If ws.Cells(start, 3).Value = 0 Then
      For find_value = start To i
     If ws.Cells(start, 3).Value <> 0 Then
       start = find_value
       
    Exit For
    End If
    Next find_value
    End If
     
' Calculate Change
   change = (ws.Cells(i, 6) - ws.Cells(start, 3))
    percentChange = change / ws.Cells(start, 3)
   ' start of the next stock ticker
     start = i + 1
   ' print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j) = change
                    'format the numbers as 0.00
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total

                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4 'Green
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3 'Red
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0 'No color
                    End Select
                End If
                ' reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                start = i + 1
                                     
                
            ' If ticker is still the same, add results
            Else
                total = total + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
        
        ' take the max and min and place them in a separate part in the worksheet
        
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))
        
        
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

       ' final ticker symbol for  total, greatest % of increase and decrease, and average
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    Next ws
End Sub



