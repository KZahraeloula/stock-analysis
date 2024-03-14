Sub stocks():
Dim ticker As String 'defining the ticker var
ticker = ""
Dim lastR As Long ' defining last row
Dim i, j As Long
Dim Ticker_Row As Long: Ticker_Row = 1
Dim YChange, OPrice, Cprice, percentage_change, temp As Double
Dim VSTOCK As Long
OPrice = 0
Cprice = 0
YChange = 0
percentage_change = 0
j = 2
'VSTOCK = 0

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
        ' Hamza Cells(Ticker_Row, 12).Value = Cells(j, 12).Value + Cells(j, 7).Value
     'Cells(Ticker_Row, 13).Value = j
     'Cells(Ticker_Row, 14).Value = i
     Cells(Ticker_Row, 15).Value = OPrice
     Cells(Ticker_Row, 16).Value = Cprice
       'conditional formating depending on value of yearly change
         If Cells(Ticker_Row, 10).Value <= 0 Then Cells(Ticker_Row, 10).Interior.ColorIndex = 3
         ElseIf Cells(Ticker_Row, 10).Value > 0 Then Cells(Ticker_Row, 10).Interior.ColorIndex = 4
         End If
   

 Next i
 
End Sub
