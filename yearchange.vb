Sub stocks():
Dim ticker As String 'defining the ticker var
ticker = ""
Dim lastR As Long ' defining last row
Dim i, j As Long
Dim Ticker_Row As Long: Ticker_Row = 1
'Dim first_row As Long: first_row = 2
Dim YChange, OPrice, Cprice As Double
OPrice = 0
Cprice = 0
YChange = 0
j = 2
'finding lastrow number

 lastR = Cells(Rows.Count, 1).End(xlUp).Row
 'MsgBox (" last row is number" & lastR) test
 
 'column hearders
 Range("I1").Value = "Ticker"
 Range("J1").Value = "Year Change"
 
 'loop for ticker
 For i = 2 To lastR  'going througt all the rows
   OPrice = Cells(j, 3).Value
 
   
  If Cells(i + 1, 1).Value <> Cells(i, 1) Then  ' testing if the ticker changed from one row to another
  ticker = Cells(i, 1).Value                    'giving ticker the value
  Ticker_Row = Ticker_Row + 1                   'going down a row
  Cells(Ticker_Row, 9).Value = ticker           'writing in the colunm I
  Cprice = Cells(i, 6)
  'Cells(Ticker_Row, 10).Value = Cells(j, 3).Value
  j = i + 1
  'MsgBox (OPrice & "   " & Cprice)
  'Cells(Ticker_Row, 11).Value = Cprice
  YChange = Cprice - OPrice
  Cells(Ticker_Row, 10).Value = YChange
  'else if oprice
  
  
  End If
  