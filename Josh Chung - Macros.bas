Attribute VB_Name = "Module1"
Sub EasyStock():
'Creates headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Annual Volume"

'Defines variables
Dim Ticker As String
Dim Volume As Double
    Volume = 0
Dim VolumeSum As Double
    VolumeSum = 0
Dim y As Long
Dim CounterRow As Double
Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
'Loops through rows
For y = 2 To LastRow
    Ticker = Cells(y, 1).Value
    Volume = Cells(y, 7).Value / 100 'Divides values by 100
    If Ticker = Cells(y + 1, 1).Value Then
        VolumeSum = VolumeSum + Volume
    ElseIf Ticker <> Cells(y + 1, 1).Value Then
        VolumeSum = VolumeSum + Volume
        CounterRow = Cells(Rows.Count, 9).End(xlUp).Row + 1
        Cells(CounterRow, 9).Value = Ticker
        Cells(CounterRow, 10).Value = VolumeSum * 100 'Multiplies values by 100
        VolumeSum = 0
    End If
Next y
End Sub
Sub ModerateStock():
'Creates headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Defines variables
Dim Ticker As String
Dim Volume As Double, VolumeSum As Double
Dim CounterRow As Double
Dim OpenPrice As Double, ClosePrice As Double
Dim PriceDiff As Double, PricePercent As Double
Dim y As Long, LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
'Loops through rows
For y = 2 To LastRow
    Ticker = Cells(y, 1).Value
    Volume = Cells(y, 7).Value / 100 'Divided by 100 to make numbers more managable
    If Ticker <> Cells(y - 1, 1).Value Then 'If the previous line doesn't match the current line
        VolumeSum = VolumeSum + Volume
        OpenPrice = Cells(y, 3).Value
    ElseIf Ticker = Cells(y + 1, 1).Value Then 'If the next line matches the current line
        VolumeSum = VolumeSum + Volume
    ElseIf Ticker <> Cells(y + 1, 1).Value Then 'If the next line doesnt match the current line
        VolumeSum = VolumeSum + Volume
        CounterRow = Cells(Rows.Count, 9).End(xlUp).Row + 1
        
        'Fills in stock ticker
        Cells(CounterRow, 9).Value = Ticker
        
        'Calculates price difference, stores values in cells and adds color
        ClosePrice = Cells(y, 6).Value
        PriceDiff = ClosePrice - OpenPrice
        Cells(CounterRow, 10).Value = PriceDiff
            Cells(CounterRow, 10).NumberFormat = "0.0000000"
            If PriceDiff > 0 Then
                Cells(CounterRow, 10).Interior.ColorIndex = 4
            ElseIf PriceDiff <= 0 Then
                Cells(CounterRow, 10).Interior.ColorIndex = 3
            End If
            
        'Calculates price percentage and stores values in cells
        If OpenPrice = 0 Then
            PricePercent = PriceDiff
        Else
        PricePercent = PriceDiff / OpenPrice
        Cells(CounterRow, 11).Value = PricePercent
            Cells(CounterRow, 11).NumberFormat = "0.00%"
        End If
        
        'Fills in total stock volume
        Cells(CounterRow, 12).Value = VolumeSum * 100 'Multiplied values by 100 to restore
        
        'Resets variable to zero for next loop
        VolumeSum = 0
    End If
Next y
End Sub
Sub HardStock():
'Creates headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(2, 14).Value = "Greatest % Increase"
Cells(3, 14).Value = "Grestest % Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"

'Defines variables
Dim Ticker As String
Dim Volume As Double, VolumeSum As Double
Dim CounterRow As Double
Dim OpenPrice As Double, ClosePrice As Double
Dim PriceDiff As Double, PricePercent As Double
Dim y As Long, LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
    
'Loops through rows
For y = 2 To LastRow
    Ticker = Cells(y, 1).Value
    Volume = Cells(y, 7).Value / 100 'Divided by 100 to make numbers more managable
    If Ticker <> Cells(y - 1, 1).Value Then 'If the previous line doesn't match the current line
        VolumeSum = VolumeSum + Volume
        OpenPrice = Cells(y, 3).Value
    ElseIf Ticker = Cells(y + 1, 1).Value Then 'If the next line matches the current line
        VolumeSum = VolumeSum + Volume
    ElseIf Ticker <> Cells(y + 1, 1).Value Then 'If the next line doesnt match the current line
        VolumeSum = VolumeSum + Volume
        CounterRow = Cells(Rows.Count, 9).End(xlUp).Row + 1
        
        'Fills in stock ticker
        Cells(CounterRow, 9).Value = Ticker
        
        'Calculates price difference, stores values in cells and adds color
        ClosePrice = Cells(y, 6).Value
        PriceDiff = ClosePrice - OpenPrice
        Cells(CounterRow, 10).Value = PriceDiff
            Cells(CounterRow, 10).NumberFormat = "0.0000000"
            If PriceDiff > 0 Then
                Cells(CounterRow, 10).Interior.ColorIndex = 4
            ElseIf PriceDiff <= 0 Then
                Cells(CounterRow, 10).Interior.ColorIndex = 3
            End If
            
        'Calculates price percentage and stores values in cells
        If OpenPrice = 0 Then
            PricePercent = PriceDiff
        Else
        PricePercent = PriceDiff / OpenPrice
        Cells(CounterRow, 11).Value = PricePercent
            Cells(CounterRow, 11).NumberFormat = "0.00%"
        End If
            
        'Fills in total stock volume
        Cells(CounterRow, 12).Value = VolumeSum * 100 'Multiplied values by 100 to restore
        
        'Resets variable to zero for next loop
        VolumeSum = 0
    End If
Next y

'Loop through results for final
Dim DataRow As Integer
Dim j As Integer
Dim PercentIncrease As Double
Dim PercentDecrease As Double
Dim HighestTotal As Double
Dim PercentHold As Double
Dim TotalHold As Double
Dim HighestTotalStock As String
Dim PercentIncreaseStock As String
Dim PercentDecreaseStock As String
DataRow = Cells(Rows.Count, 9).End(xlUp).Row + 1
HigestTotal = 0
PercentIncrease = 0
PercentDecrease = 0

For j = 2 To DataRow
        PercentHold = Cells(j, 11).Value
        TotalHold = Cells(j, 12).Value
        If PercentHold > PercentIncrease Then
            PercentIncrease = PercentHold
            PercentIncreaseStock = Cells(j, 9).Value
        End If
        If PercentHold < PercentDecrease Then
            PercentDecrease = PercentHold
            PercentDecreaseStock = Cells(j, 9).Value
        End If
        If TotalHold > HighestTotal Then
            HighestTotal = TotalHold
            HighestTotalStock = Cells(j, 9).Value
        End If
Next j

'Print final results
Cells(2, 16).Value = PercentIncrease
Cells(2, 16).NumberFormat = "%####.##"
Cells(2, 15).Value = PercentIncreaseStock
Cells(3, 16).Value = PercentDecrease
Cells(3, 16).NumberFormat = "%####.##"
Cells(3, 15).Value = PercentDecreaseStock
Cells(4, 16).Value = HighestTotal
Cells(4, 15).Value = HighestTotalStock
End Sub
