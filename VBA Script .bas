Attribute VB_Name = "Module1"
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim startRow As Long

    ' Loop through all sheets (representing each quarter)
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row of data in the current sheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize headers for the output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables for greatest values
        greatestIncrease = -1E+30 ' Set to very low value initially
        greatestDecrease = 1E+30  ' Set to very high value initially
        greatestVolume = 0
        greatestIncreaseTicker = "" ' Initialize as empty string
        greatestDecreaseTicker = ""  ' Initialize as empty string
        greatestVolumeTicker = ""     ' Initialize as empty string
        startRow = 2 ' Assuming headers are in row 1
        
        ' Loop through each row of data
        For i = startRow To lastRow
            ticker = ws.Cells(i, 1).Value ' Ticker Symbol (column A)
            ' Check if the cell is numeric before assigning
            If IsNumeric(ws.Cells(i, 3).Value) Then
                openPrice = ws.Cells(i, 3).Value  ' Open Price (column C)
            Else
                openPrice = 0 ' or handle it as needed
            End If
            
            If IsNumeric(ws.Cells(i, 6).Value) Then
                closePrice = ws.Cells(i, 6).Value ' Close Price (column F)
            Else
                closePrice = 0 ' or handle it as needed
            End If
            
            If IsNumeric(ws.Cells(i, 7).Value) Then
                totalVolume = ws.Cells(i, 7).Value ' Volume (column G)
            Else
                totalVolume = 0 ' or handle it as needed
            End If
            
            ' Calculate quarterly change and percent change
            quarterlyChange = closePrice - openPrice
            
            ' Avoid division by zero
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0 ' or handle it as needed
            End If
           Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim startRow As Long

    ' Loop through all sheets (representing each quarter)
    For Each ws In ThisWorkbook.Sheets
        ' Find the last row of data in the current sheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize headers for the output columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change ($)"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Initialize variables for greatest values
        greatestIncrease = -1E+30 ' Set to very low value initially
        greatestDecrease = 1E+30  ' Set to very high value initially
        greatestVolume = 0
        greatestIncreaseTicker = "" ' Initialize as empty string
        greatestDecreaseTicker = ""  ' Initialize as empty string
        greatestVolumeTicker = ""     ' Initialize as empty string
        startRow = 2 ' Assuming headers are in row 1
        
        ' Loop through each row of data
        For i = startRow To lastRow
            ticker = ws.Cells(i, 1).Value ' Ticker Symbol (column A)
            ' Check if the cell is numeric before assigning
            If IsNumeric(ws.Cells(i, 3).Value) Then
                openPrice = ws.Cells(i, 3).Value  ' Open Price (column C)
            Else
                openPrice = 0 ' or handle it as needed
            End If
            
            If IsNumeric(ws.Cells(i, 6).Value) Then
                closePrice = ws.Cells(i, 6).Value ' Close Price (column F)
            Else
                closePrice = 0 ' or handle it as needed
            End If
            
            If IsNumeric(ws.Cells(i, 7).Value) Then
                totalVolume = ws.Cells(i, 7).Value ' Volume (column G)
            Else
                totalVolume = 0 ' or handle it as needed
            End If
            
            ' Calculate quarterly change and percent change
            quarterlyChange = closePrice - openPrice
            
            ' Avoid division by zero
            If openPrice <> 0 Then
                percentChange = (quarterlyChange / openPrice) * 100
            Else
                percentChange = 0 ' or handle it as needed
            End If
            
            ' Output results into columns 9, 10, 11, 12 (ticker, volume, quarterly change, percent change)
            ws.Cells(i, 9).Value = ticker ' Ticker Symbol (Column 9)
            ws.Cells(i, 10).Value = quarterlyChange ' Percent Change (Column 10)
            ws.Cells(i, 11).Value = percentChange ' Quarterly Change ($) (Column 11)
            ws.Cells(i, 11).NumberFormat = "0.00%" ' Format as percentage with 2 decimal places
            ws.Cells(i, 12).Value = totalVolume ' Total Stock Volume (Column 12)
            
            ' Find the stock with the greatest % increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Place the results for the greatest values in row 14
        ws.Cells(14, 2).Value = greatestIncreaseTicker ' Greatest % Increase (Column 14,2)
        ws.Cells(14, 3).Value = greatestDecreaseTicker ' Greatest % Decrease (Column 14,3)
        ws.Cells(14, 4).Value = greatestVolumeTicker ' Greatest Total Volume (Column 14,4)
        
        ' Apply Conditional Formatting for Quarterly Change (Column 11)
        With ws.Range("J2:J" & lastRow) ' Column 11: Quarterly Change
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
         End With
            
        With ws.Range("K2:K" & lastRow) ' Column 11: Quarterly Change
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With
        
     Next ws
     
End Sub
            ' Output results into columns 9, 10, 11, 12 (ticker, volume, quarterly change, percent change)
            ws.Cells(i, 9).Value = ticker ' Ticker Symbol (Column 9)
            ws.Cells(i, 10).Value = quarterlyChange ' Percent Change (Column 10)
            ws.Cells(i, 11).Value = percentChange ' Quarterly Change ($) (Column 11)
            ws.Cells(i, 11).NumberFormat = "0.00%" ' Format as percentage with 2 decimal places
            ws.Cells(i, 12).Value = totalVolume ' Total Stock Volume (Column 12)
            
            ' Find the stock with the greatest % increase, decrease, and volume
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
        Next i
        
        ' Place the results for the greatest values in row 14
        ws.Cells(15, 2).Value = greatestIncreaseTicker ' Greatest % Increase (Column 14,2)
        ws.Cells(15, 3).Value = greatestDecreaseTicker ' Greatest % Decrease (Column 14,3)
        ws.Cells(15, 4).Value = greatestVolumeTicker ' Greatest Total Volume (Column 14,4)
        
        ' Apply Conditional Formatting for Quarterly Change (Column 11)
        With ws.Range("J2:J" & lastRow) ' Column 11: Quarterly Change
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
         End With
            
        With ws.Range("K2:K" & lastRow) ' Column 11: Quarterly Change
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .FormatConditions(1).Interior.Color = RGB(0, 255, 0) ' Green for positive change
            .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .FormatConditions(2).Interior.Color = RGB(255, 0, 0) ' Red for negative change
        End With
        
     Next ws
     
End Sub
