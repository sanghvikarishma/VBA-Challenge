Sub Multi_year_stock_data()
    
    Dim ws As Worksheet
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim summaryTableRow As Long
    Dim yearBeginRow As Long
    Dim lastvalue As Long
    
    For Each ws In Worksheets
        
        ' Initialize variables
        ticker = 0
        openingPrice = 2
        closingPrice = 0
        yearlyChange = 0
        percentChange = 0
        totalstockVolume = 0
        summaryTableRow = 2

ws.Range("i1").Value = "ticker"
ws.Range("j1").Value = "yearly change"
ws.Range("k1").Value = "percent change"
ws.Range("l1").Value = "total stock volume"
ws.Range("o2").Value = "greatest percent increase"
ws.Range("o3").Value = "greatest percent decrease"
ws.Range("o4").Value = "greatest total volume"
ws.Range("p1").Value = "ticker"
ws.Range("q1").Value = "value"



  ' Find the last row of the data
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     openingPrice = ws.Cells(2, 3).Value
     lastvalue = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
    ' Loop through the rows of the data
      For i = 2 To lastRow
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
tickername = ws.Cells(i, 1).Value
tickertotal = tickertotal + ws.Cells(i, 7).Value
counter = WorksheetFunction.CountIf(ws.Range("A:A"), tickername)

openingPrice = ws.Cells(i - counter + 1, 3).Value
closingPrice = ws.Cells(i, 6)
yearlyChange = closingPrice - openingPrice
percentChange = yearlyChange / openingPrice


If yearlyChange < 0 Then
ws.Range("J" & summaryTableRow).Interior.ColorIndex = 3

Else
ws.Range("J" & summaryTableRow).Interior.ColorIndex = 4

End If
          
      
'print all results
ws.Range("i" & summaryTableRow).Value = tickername
ws.Range("j" & summaryTableRow).Value = yearlyChange
ws.Range("k" & summaryTableRow).Value = percentChange
ws.Range("k" & summaryTableRow).NumberFormat = "0.00%"
ws.Range("l" & summaryTableRow).Value = tickertotal

'end loop and reset variables
summaryTableRow = summaryTableRow + 1
tickertotal = 0



Else
'continue looping
tickertotal = tickertotal + ws.Cells(i, 7).Value
End If

 
 Next i
 
 
'find greatest percent increase
maxvalue = ws.Application.Max(ws.Range("k:k"))
ws.Range("q2").Value = maxvalue
ws.Range("q2").NumberFormat = "0.00%"

For k = 2 To lastvalue
If ws.Range("k" & k).Value = maxvalue Then
ws.Range("p2").Value = ws.Range("I" & k).Value
End If

Next k


' find greatest percent decrease
minvalue = ws.Application.Min(ws.Range("k:k"))
ws.Range("q3").Value = minvalue
ws.Range("q3").NumberFormat = "0.00%"

For k = 2 To lastvalue
If ws.Range("k" & k).Value = minvalue Then
ws.Range("P3").Value = ws.Range("I" & k).Value
End If

Next k

'find greatest total volume
gtv = ws.Application.Max(ws.Range("l:l"))
ws.Range("q4").Value = gtv

For k = 2 To lastvalue
If ws.Range("l" & k).Value = gtv Then
ws.Range("p4").Value = ws.Range("i" & k).Value
End If



Next k

Next ws










End Sub
