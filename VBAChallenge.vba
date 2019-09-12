Sub VBAHomework()

'defining the variables
Dim Ticker As String
Dim TickerTotal As Variant
Dim SummaryTableRow As Long
Dim fp As Single 'final price
Dim IP As Single 'initial price
Dim yp As Single 'yearly change
Dim pc As Single 'percent change
Dim i As Long     'contador
Dim lastrow As Double 'ultima fila

'Setting the titles in each sheet
    range("j1").Value = "Ticker"
    range("k1").Value = "Yearly Change"
    range("l1").Value = "Percent Change"
    range("m1").Value = "Total Stock Volume"

'set initial values
i = 2  'definir i=2 para establecer el primer valor del contador en las filas
SummaryTableRow = 2
TickerTotal = 0

'asignar ip a celda C2
IP = Cells(i, 3).Value


lastrow = range("a" & Rows.Count).End(xlUp).Row


'If a2 = a Then
'CIP = cell(3, 2).Value
'End If

'loop trough the rows for the Tricker number
For i = 2 To lastrow

'check the cells with the same Ticker number
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
          Ticker = Cells(i, 1).Value
              TickerTotal = TickerTotal + Cells(i, 7).Value
              range("j" & SummaryTableRow).Value = Ticker
              range("m" & SummaryTableRow).Value = TickerTotal
                
'setting price of the first day vs last day price
              fp = Cells(i, 6)
              yp = fp - IP
              range("k" & SummaryTableRow).Value = yp
                If IP <> 0 Then
                     pc = (fp / IP) - 1
                Else
                     pc = 1
                End If
              range("L" & SummaryTableRow).Value = pc
' colors positives green and negatives red
                If yp > 0 Then
                    range("K" & SummaryTableRow).Interior.ColorIndex = 4
                ElseIf yp < 0 Then
                    range("K" & SummaryTableRow).Interior.ColorIndex = 3
                Else
                    range("K" & SummaryTableRow).Interior.ColorIndex = 0
                End If
                         
              IP = 0
              IP = Cells(i + 1, 3)
              TickerTotal = 0
              SummaryTableRow = SummaryTableRow + 1
                
    Else
      TickerTotal = TickerTotal + Cells(i, 7).Value
    
     
   End If
Next i

'setting the others variabes for the Challenge
Dim MaxTicker, MinTicker, MaxVT  As String
Dim MaxPerc, MinPerc, MaxVol As Double
Dim lastrow2 As Double
Dim J As Double

'Initial Instructions, allocate firs values once before running the loop
'Set best performer equal to the first stock
MaxPerc = Cells(2, 12)
MinPerc = Cells(2, 12)
MaxVol = Cells(2, 13)

'Print and allocate the ID Columns for the Summary Table
    
Cells(2, 16).Value = "Max Increase (%)"
Cells(3, 16).Value = "Min Increase (%)"
Cells(4, 16).Value = "Highest Volume"
Cells(1, 17).Value = "Ticker ID"
Cells(1, 18).Value = "Volume Value"

'Instruction, last row

lastrow2 = Cells(Rows.Count, 10).End(xlUp).Row

        
'Loop to search through summary table
        For J = 2 To lastrow2

                'Instruction, concetrates ticker and Percentage HIGHEST values
            If Cells(J, 12).Value > MaxPerc Then
                MaxPerc = Cells(J, 12)
                MaxTicker = Cells(J, 10).Value
            End If

                'Instruction, concentrates ticker and Percentage LEAST values
            If Cells(J, 12).Value < MinPerc Then
                MinPerc = Cells(J, 12)
                MinTicker = Cells(J, 10)
                
            End If

                'Conditional to determine stock with the greatest volume traded
            If Cells(J, 13) > MaxVol Then
                MaxVol = Cells(J, 13).Value
                MaxVT = Cells(J, 10)
                
            End If

        Next J

'Putting the information in a table
        Cells(2, 17).Value = MaxTicker
        Cells(2, 18).Value = MaxPerc
        Cells(2, 18).NumberFormat = "0.00%"
        Cells(3, 17).Value = MinTicker
        Cells(3, 18).Value = MinPerc
        Cells(3, 18).NumberFormat = "0.00%"
        Cells(4, 17).Value = MaxVT
        Cells(4, 18).Value = MaxVol

End Sub
