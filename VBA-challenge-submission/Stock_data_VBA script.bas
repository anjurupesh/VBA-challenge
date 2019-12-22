Attribute VB_Name = "Module1"
Sub StockData()
  
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        '
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

       
        Cells(1, "J").Value = "Ticker"
        Cells(1, "K").Value = "Yearly Change"
        Cells(1, "L").Value = "Percent Change"
        Cells(1, "M").Value = "Total Stock Volume"
       
        Dim Opening_Price As Double
        Dim Closeing_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
      
        Open_Price = Cells(2, Column + 2).Value
         ' Loop through all ticker symbol
        
        For i = 2 To LastRow
        
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, Column + 9).Value = Ticker_Name
                
                Close_Price = Cells(i, Column + 5).Value
               
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 10).Value = Yearly_Change
               
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 11).Value = Percent_Change
                    Cells(Row, Column + 11).NumberFormat = "0.00%"
                End If
               
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(Row, Column + 12).Value = Volume
                
                Row = Row + 1
               
                Open_Price = Cells(i + 1, Column + 2)
                
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Determine the Last Row of Yearly Change per WS
        YCLastRow = WS.Cells(Rows.Count, Column + 9).End(xlUp).Row
        
        For j = 2 To YCLastRow
            If (Cells(j, Column + 10).Value > 0 Or Cells(j, Column + 10).Value = 0) Then
                Cells(j, Column + 10).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 10).Value < 0 Then
                Cells(j, Column + 10).Interior.ColorIndex = 3
            End If
        Next j
      
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
      
        For Z = 2 To YCLastRow
            If Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 9).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 11).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Min(WS.Range("L2:L" & YCLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 9).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 11).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 12).Value = Application.WorksheetFunction.Max(WS.Range("M2:M" & YCLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 9).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 12).Value
            End If
        Next Z
        
    Next WS
        
End Sub


