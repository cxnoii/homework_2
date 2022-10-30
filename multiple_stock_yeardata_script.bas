Attribute VB_Name = "Module1"
Sub Wall_st()

'declaring variables
Dim lastRow As Double
Dim openPriceCount As Double
Dim volume As Double

Dim tableRow As Integer
Dim stockName As String

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

'variable that increments to find index of openPrice (i-openPriceCount) returns index
openPriceCount = 0
volume = 0
tableRow = 2

'this loop creates table & populates
For I = 2 To lastRow
    If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        stockName = Cells(I, 1).Value
        volume = volume + Cells(I, 7).Value
        'prints ticker
        Cells(tableRow, 9).Value = Cells(I, 1).Value
        
        'prints volume
        Cells(tableRow, 12).Value = volume
        
        'prints yearly change, openPriceCount incremented everytime no string difference was detected
        Cells(tableRow, 10).Value = Cells(I, 6).Value - Cells(I - openPriceCount, 3)
        
        'print %change
        Cells(tableRow, 11).Value = ((Cells(I, 6).Value - Cells(I - openPriceCount, 3)) / (Cells(I - openPriceCount, 3))) * 100
        
        'increment row&reset
        tableRow = tableRow + 1
        volume = 0
        
    Else
    volume = Cells(I, 7).Value + volume
    openPriceCount = openPriceCount + 1
    
    End If
    
    Next I
    
'this loop applies conditional formatting
For I = 2 To (tableRow)
    If Cells(I, 10).Value > 0 Then
        Cells(I, 10).Interior.ColorIndex = 4
    ElseIf Cells(I, 10).Value < 0 And Cells(I, 10).Value <> 0 Then
        Cells(I, 10).Interior.ColorIndex = 3
    End If

Next I

'declaring more variables
Dim bigInc As Double
Dim smallInc As Double
Dim bigVol As Double
Dim lastRowSummary As Integer

Dim stockname2 As String
stockname2 = ""
lastRowSummary = Cells(Rows.Count, 9).End(xlUp).Row

'greatest %increase + populate table2
For I = 2 To lastRowSummary
    If Cells(I, 11).Value > Cells(I + 1, 11).Value And Cells(I, 11).Value > bigInc Then
    bigInc = Cells(I, 11).Value
    stockname2 = Cells(I, 9).Value
    End If
    Next I
Cells(2, 15).Value = stockname2
Cells(2, 16).Value = bigInc

'greatest %decrease + populate table2
For I = 2 To lastRowSummary
    If Cells(I, 11).Value < Cells(I + 1, 11).Value And Cells(I, 11).Value < smallInc Then
    smallInc = Cells(I, 11).Value
    stockname2 = Cells(I, 9).Value
    End If
    Next I
Cells(3, 15).Value = stockname2
Cells(3, 16).Value = smallInc

'greatest volume + populate table2
For I = 2 To lastRowSummary
    If Cells(I, 12).Value > Cells(I + 1, 12).Value And Cells(I, 12).Value > bigVol Then
    bigVol = Cells(I, 12).Value
    stockname2 = Cells(I, 9).Value
    End If
    Next I
Cells(4, 15).Value = stockname2
Cells(4, 16).Value = bigVol

'printing row headers for table1
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Yearly Change"
Cells(1, 12).Value = "Stock Volume"
       
'printing row/column headers for table2
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Greatest % Decrease"
Cells(4, 14).Value = "Greatest Volume"

End Sub


