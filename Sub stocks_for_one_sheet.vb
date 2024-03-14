Sub stocks():
Dim ticker As String 'defining the ticker var
ticker = ""
Dim lastR, lastK As Long ' defining last row
Dim i, j As Long 'i to find last row in column A ,J to find first row for each ticker , K to find last row of column K
Dim Ticker_Row As Long: Ticker_Row = 1
Dim YChange, OPrice, Cprice, percentage_change, temp As Double
Dim VSTOCK As Long
Dim GPI, GPD As Double
Dim GVol, TotVol As Double
Dim k As Long ' k to run through the column k
Dim pincrease, pdecrease As Double
Dim tickerstat, gticker, sticher, vticker As String 'tickerstat counter for ticker value, gticker greatest ticker, smallest ticher, vticker for volume ticker


OPrice = 0
Cprice = 0
YChange = 0
percentage_change = 0
j = 2
GPI = 0
GPD = 0
GVol = 0

lastR = Cells(Rows.Count, 1).End(xlUp).Row 'finding lastrow number
'MsgBox (" last row is number" & lastR) test
'column hearders
Range("I1").Value = "Ticker"
Range("J1").Value = "Year Change"
Range("k1").Value = "Percentage change"
Range("L1").Value = "Stock Volume"
Range("M1").Value = "J"
Range("N1").Value = "I"
Range("o1").Value = "OPrice"
Range("p1").Value = "CPrice"

Range("R2").Value = "Greatest % increase"
Range("R3").Value = "Greatest % decrease"
Range("R4").Value = "Greatest total volume"
Range("s1").Value = " ticker"
Range("t1").Value = " Value"
'loop for ticker
For i = 2 To lastR  'going througt all the rows
   'OPrice = Cells(j, 3).Value
   If Cells(i + 1, 1).Value <> Cells(i, 1) Then  ' testing if the ticker changed from one row to another
   ticker = Cells(i, 1).Value                     'giving ticker the value
   Ticker_Row = Ticker_Row + 1                   'going down a row
   Cells(Ticker_Row, 9).Value = ticker           'writing in the colunm I
   OPrice = Cells(j, 3).Value
   Cprice = Cells(i, 6)
   Cells(Ticker_Row, 13).Value = j
    Cells(Ticker_Row, 14).Value = i
     Cells(Ticker_Row, 12).Value = Application.Sum(Range(Cells(j, 7), Cells(i, 7)))
      'Cells(Ticker_Row, 10).Value = Cells(j, 3).Value
    'End If
     j = i + 1
    
       'VSTOCK = Cells(j, 7).Value
       'MsgBox (OPrice & "   " & Cprice)
       'Cells(Ticker_Row, 11).Value = Cprice
    YChange = Cprice - OPrice
    Cells(Ticker_Row, 10).Value = YChange
    percentage_change = (YChange / OPrice)
    Cells(Ticker_Row, 11).Value = percentage_change
    Cells(Ticker_Row, 11).NumberFormat = "0.00 %"   'formating the percent change
    'Cells(Tick_Row, 12).Value = Function.Sum(Range(Cells(j, 7), Cells(i, 7))
    'Cells(Tick_Row, 12).Value = Application.Sum(Range(Cells(i, 7), Cells(j, 7)))
    'Cells(Ticker_Row, 12).Value = Application.Sum(Range(Cells(j, 7), Cells(i, 7)))
    'Cells(Ticker_Row, 13).Value = j
    'Cells(Ticker_Row, 14).Value = i
    Cells(Ticker_Row, 15).Value = OPrice
    Cells(Ticker_Row, 16).Value = Cprice
      'conditional formating depending on value of yearly change
        If Cells(Ticker_Row, 10).Value <= 0 Then Cells(Ticker_Row, 10).Interior.ColorIndex = 3
        ElseIf Cells(Ticker_Row, 10).Value > 0 Then Cells(Ticker_Row, 10).Interior.ColorIndex = 4
        End If
 
Next i
lastK = Cells(Rows.Count, 11).End(xlUp).Row 'finding lastrow number in column K
'Range("q1").Value = lastK 'test k
'loop for greatest percent increase
k = 2
gticker = ""
For k = 2 To lastK
tickerstat = Cells(k, 9).Value
pincrease = Cells(k, 11).Value
pdecrease = Cells(k, 11).Value
GVol = Cells(k, 12).Value
   If pincrease > GPI Then
  GPI = pincrease
  gticker = tickerstat
   End If
 Next k
 'this gives me the correcte numbers but no way i can get the ticker
 'GPI = Application.Max(Range(Cells(2, 11), Cells(lastK, 11)))
 'GPD = Application.Min(Range(Cells(2, 11), Cells(lastK, 11)))
 'GVol = Application.Max(Range(Cells(2, 12), Cells(lastK, 12)))
Cells(2, 20).Value = GPI
Cells(2, 20).NumberFormat = "0.00 %"
Cells(2, 19).Value = gticker
'k = 2
'loop for greatest percet decrease
sticker = ""
For k = 2 To lastK
tickerstat = Cells(k, 9).Value
'pdecrease = Cells(k, 11).Value
pdecrease = Cells(k, 11).Value
'GVol = Cells(k, 12).Value
  If pdecrease < GPD Then
  GPD = pdecrease
  sticker = tickerstat
  End If
 Next k
Cells(3, 20).Value = GPD
Cells(3, 20).NumberFormat = "0.00 %"
Cells(3, 19).Value = sticker

'loop for greatest volum
vticker = ""
For k = 2 To lastK
tickerstat = Cells(k, 9).Value
'pdecrease = Cells(k, 11).Value
'pdecrease = Cells(k, 11).Value
TotVol = Cells(k, 12).Value
  If TotVol > GVol Then
  GVol = TotVol
  vticker = tickerstat
  End If
 Next k
Cells(4, 20).Value = GVol
'Cells(3, 20).NumberFormat = "0.00 %"
Cells(4, 19).Value = vticker


'Cells(4, 20).Value = GVol

End Sub

