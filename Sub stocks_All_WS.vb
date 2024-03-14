Sub stocks():

    
Dim WorksheetName As String
        
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

For Each Ws In Worksheets
WorksheetName = Ws.Name
OPrice = 0
Cprice = 0
YChange = 0
percentage_change = 0
j = 2
GPI = 0
GPD = 0
GVol = 0

lastR = Ws.Cells(Rows.Count, 1).End(xlUp).Row 'finding lastrow number
'MsgBox (" last row is number" & lastR) test
'column hearders
Ws.Range("I1").Value = "Ticker"
Ws.Range("J1").Value = "Year Change"
Ws.Range("k1").Value = "Percentage change"
Ws.Range("L1").Value = "Stock Volume"
Ws.Range("M1").Value = "J"
Ws.Range("N1").Value = "I"
Ws.Range("o1").Value = "OPrice"
Ws.Range("p1").Value = "CPrice"

Ws.Range("R2").Value = "Greatest % increase"
Ws.Range("R3").Value = "Greatest % decrease"
Ws.Range("R4").Value = "Greatest total volume"
Ws.Range("s1").Value = " Ticker"
Ws.Range("t1").Value = " Value"
Ticker_Row = 1
'loop for ticker
For i = 2 To lastR  'going througt all the rows
   If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1) Then  ' testing if the ticker changed from one row to another
   ticker = Ws.Cells(i, 1).Value                     'giving ticker the value
   Ticker_Row = Ticker_Row + 1                   'going down a row
   Ws.Cells(Ticker_Row, 9).Value = ticker           'writing in the colunm I
   OPrice = Ws.Cells(j, 3).Value
   Cprice = Ws.Cells(i, 6).Value
   Ws.Cells(Ticker_Row, 13).Value = j
   Ws.Cells(Ticker_Row, 14).Value = i
   Ws.Cells(Ticker_Row, 12).Value = Application.Sum(Range(Ws.Cells(j, 7), Ws.Cells(i, 7)))
     
    j = i + 1

    'MsgBox (OPrice & "   " & Cprice)
    YChange = Cprice - OPrice
    Ws.Cells(Ticker_Row, 10).Value = YChange
    percentage_change = (YChange / OPrice)
    Ws.Cells(Ticker_Row, 11).Value = percentage_change
    Ws.Cells(Ticker_Row, 11).NumberFormat = "0.00 %"   'formating the percent change
    'Cells(Tick_Row, 12).Value = Function.Sum(Range(Cells(j, 7), Cells(i, 7))
    'Cells(Tick_Row, 12).Value = Application.Sum(Range(Cells(i, 7), Cells(j, 7)))
    'Cells(Ticker_Row, 12).Value = Application.Sum(Range(Cells(j, 7), Cells(i, 7)))
    'Cells(Ticker_Row, 13).Value = j
    'Cells(Ticker_Row, 14).Value = i
    Ws.Cells(Ticker_Row, 15).Value = OPrice
    Ws.Cells(Ticker_Row, 16).Value = Cprice
      'conditional formating depending on value of yearly change
        If Ws.Cells(Ticker_Row, 10).Value <= 0 Then Ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 3
        ElseIf Ws.Cells(Ticker_Row, 10).Value > 0 Then Ws.Cells(Ticker_Row, 10).Interior.ColorIndex = 4
        End If
 
Next i
lastK = Ws.Cells(Rows.Count, 11).End(xlUp).Row 'finding lastrow number in column K
'Range("q1").Value = lastK 'test k
'loop for greatest percent increase
k = 2
gticker = ""
For k = 2 To lastK
tickerstat = Ws.Cells(k, 9).Value
pincrease = Ws.Cells(k, 11).Value
pdecrease = Ws.Cells(k, 11).Value
GVol = Ws.Cells(k, 12).Value
   If pincrease > GPI Then
  GPI = pincrease
  gticker = tickerstat
   End If
 Next k
 'this gives me the correcte numbers but no way I can get the ticker
 'GPI = Application.Max(Range(Cells(2, 11), Cells(lastK, 11)))
 'GPD = Application.Min(Range(Cells(2, 11), Cells(lastK, 11)))
 'GVol = Application.Max(Range(Cells(2, 12), Cells(lastK, 12)))
Ws.Cells(2, 20).Value = GPI
Ws.Cells(2, 20).NumberFormat = "0.00 %"
Ws.Cells(2, 19).Value = gticker
'k = 2
'loop for greatest percet decrease
sticker = ""
For k = 2 To lastK
tickerstat = Ws.Cells(k, 9).Value
pdecrease = Ws.Cells(k, 11).Value
  If pdecrease < GPD Then
  GPD = pdecrease
  sticker = tickerstat
  End If
 Next k
Ws.Cells(3, 20).Value = GPD
Ws.Cells(3, 20).NumberFormat = "0.00 %"
Ws.Cells(3, 19).Value = sticker

'loop for greatest volume
vticker = ""
For k = 2 To lastK
tickerstat = Ws.Cells(k, 9).Value
TotVol = Ws.Cells(k, 12).Value
  If TotVol > GVol Then
  GVol = TotVol
  vticker = tickerstat
  End If
 Next k
Ws.Cells(4, 20).Value = GVol
Ws.Cells(4, 19).Value = vticker

Next Ws
End Sub


