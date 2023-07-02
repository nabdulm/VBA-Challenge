Attribute VB_Name = "Module1"
Sub PB()
Dim readName As String
Dim nextName As String
Dim groupNo As Long
Dim totalSV As Double
Dim openPrice As Double
Dim closePrice As Double
Dim maxIncP As Double
Dim maxDecP As Double
Dim maxTV As Double
Dim maxIncT As String
Dim maxDecT As String
Dim maxTVT As String
Dim i As Long
Dim lastRow As Long

Dim curSheet As Worksheet

For Each curSheet In ActiveWorkbook.Worksheets

curSheet.Cells(1, 9).Value = "Ticker"
curSheet.Cells(1, 10).Value = "Yearly Change "
curSheet.Cells(1, 11).Value = "Percent Change "
curSheet.Cells(1, 12).Value = "Total Stock Volumn"
maxIncP = 0
maxDecP = 0
maxTV = 0
groupNo = 1
totalSV = 0
openPrice = curSheet.Cells(2, 3).Value
lastRow = curSheet.Cells(curSheet.Rows.Count, 1).End(xlUp).Row

For i = 2 To lastRow
readName = curSheet.Cells(i, 1).Value
nextName = curSheet.Cells(i + 1, 1).Value

If nextName = readName Then
totalSV = totalSV + curSheet.Cells(i, 7).Value
Else
totalSV = totalSV + curSheet.Cells(i, 7).Value
closePrice = Cells(i, 6).Value
curSheet.Cells(groupNo + 1, 9).Value = readName
curSheet.Cells(groupNo + 1, 12).Value = totalSV
If totalSV > maxTV Then
maxTV = totalSV
maxTVT = readName
End If

curSheet.Cells(groupNo + 1, 10).Value = closePrice - openPrice
curSheet.Cells(groupNo + 1, 11).Value = (closePrice - openPrice) / openPrice

If closePrice - openPrice > 0 Then
curSheet.Cells(groupNo + 1, 10).Interior.Color = RGB(0, 255, 0)
If (closePrice - openPrice) / openPrice > maxIncP Then
maxIncP = (closePrice - openPrice) / openPrice
maxIncT = readName
End If
Else
curSheet.Cells(groupNo + 1, 10).Interior.Color = RGB(255, 0, 0)
If (closePrice - openPrice) / openPrice < maxDecP Then
maxDecP = (closePrice - openPrice) / openPrice
maxDecT = readNam
End If
End If

groupNo = groupNo + 1
totalSV = 0
openPrice = curSheet.Cells(i + 1, 3).Value

End If

Next i
curSheet.Range("O2").Value = "Greatest % Increase"
 curSheet.Range("O3").Value = "Greatest % Decrease"
 curSheet.Range("O4").Value = "Greatest Total Volumn"
 curSheet.Range("P1").Value = "Ticker"
 curSheet.Range("Q1").Value = "Value"
 curSheet.Range("P2").Value = maxIncT
 curSheet.Range("P4").Value = maxTVT
 curSheet.Range("Q2").Value = maxIncP
 curSheet.Range("Q3").Value = maxDecP
 curSheet.Range("K1:K" & lastRow).NumberFormat = "0.00%"
curSheet.Cells(2, 17).NumberFormat = "0.00%"
curSheet.Cells(3, 17).NumberFormat = "0.00%"
 
 Next
 
 End Sub
