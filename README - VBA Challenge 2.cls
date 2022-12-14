VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stockTicker():

 

For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"

ws.Cells(1, 10).Value = "Yearly Change"

ws.Cells(1, 11).Value = "Percent Change"

ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(1, 16).Value = "Ticker"

ws.Cells(1, 17).Value = "Value"

ws.Cells(2, 15).Value = "Greatest % Increase"

ws.Cells(3, 15).Value = "Greatest % Decrease"

ws.Cells(4, 15).Value = "Greatest Total Volume"
 

Dim Ticker As String

Dim YearlyChange As Double

Dim PercentChange As Double

Dim Volume As Double


Dim StocksOpen As Double

Dim StocksClose As Double


Dim lastrow As Double

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 

Volume = 0
  

Dim table As Double

table = 2

   
For i = 2 To lastrow
 

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then

Volume = Volume + ws.Cells(i, 7).Value

ws.Range("L" & table).Value = Volume
   

Volume = 0

   
StocksClose = ws.Cells(i, 6)
          

If StocksOpen = 0 Then

YearlyChange = 0

PercentChange = 0

Else:

YearlyChange = StocksClose - StocksOpen

PercentChange = (StocksClose - StocksOpen) / StocksOpen

End If
  

ws.Range("J" & table).Value = YearlyChange

ws.Range("K" & table).Value = PercentChange

ws.Range("K" & table).Style = "Percent"

ws.Range("K" & table).NumberFormat = "0.00%"


table = table + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then

StocksOpen = ws.Cells(i, 3)


Else: Volume = Volume + ws.Cells(i, 7).Value

End If
 

Next i


For x = 2 To lastrow


If ws.Range("J" & x).Value > 0 Then

ws.Range("J" & x).Interior.ColorIndex = 4

   
ElseIf ws.Range("J" & x).Value < 0 Then

ws.Range("J" & x).Interior.ColorIndex = 3

           
End If

Next x


ws.Range("P1").Value = "Ticker"

ws.Range("Q1").Value = "Value"


ws.Range("O2").Value = "Greatest % Increase"

ws.Range("O3").Value = "Greatest % Decrease"

ws.Range("O4").Value = "Greatest Total Volume"

   
Dim GreatestIncrease As Double

Dim GreatestDecrease As Double

Dim GreatestVolume As Double

   
GreatestIncrease = 0

GreatestDecrease = 0

GreatestVolume = 0
 

For y = 2 To lastrow
   

If ws.Cells(y, 11).Value > GreatestIncrease Then

GreatestIncrease = ws.Cells(y, 11).Value

ws.Range("Q2").Value = GreatestIncrease

ws.Range("Q2").Style = "Percent"

ws.Range("Q2").NumberFormat = "0.00%"

ws.Range("P2").Value = ws.Cells(y, 9).Value

End If

Next y
 

For Z = 2 To lastrow

       
If ws.Cells(Z, 11).Value < GreatestDecrease Then

GreatestDecrease = ws.Cells(Z, 11).Value

ws.Range("Q3").Value = GreatestDecrease

ws.Range("Q3").Style = "Percent"

ws.Range("Q3").NumberFormat = "0.00%"

ws.Range("P3").Value = ws.Cells(Z, 9).Value

 
End If
   

Next Z
  

For n = 2 To lastrow
       

If ws.Cells(n, 12).Value > GreatestVolume Then

GreatestVolume = ws.Cells(n, 12).Value

ws.Range("Q4").Value = GreatestVolume

ws.Range("P4").Value = ws.Cells(n, 9).Value
 

End If
     

Next n
       

Next ws
  

End Sub

