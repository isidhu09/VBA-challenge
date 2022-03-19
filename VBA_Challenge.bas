Attribute VB_Name = "Module1"
Sub VBA_Challenge()

'   Setting variables
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim Counter As Integer
    Dim Sum As Double
    
    For Each WS In Worksheets
'       Column Names
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
    
'       Adding counter to properly assign row values for calculations within loop and sum value for volume
        Counter = 2
        Sum = 0
        
'       Setting open price
        OpenPrice = WS.Cells(2, 3).Value
        
'       Getting row count
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row
    
'       ------------Loop BEGIN for Ticker/YearlyChange/PercentChange----------
        For i = 2 To LastRow
        
'           Starting sum for total volume by ticker
            Sum = Sum + Cells(i, 7).Value
        
            If WS.Cells(i, 1).Value <> WS.Cells(i + 1, 1).Value Then
            
'               Setting unique ticker values
                WS.Cells(Counter, 9).Value = WS.Cells(i, 1).Value
                    
'               Getting close price from same ticker
                ClosePrice = WS.Cells(i, 6).Value
                
'               Calculating and setting yearly change
                YearlyChange = ClosePrice - OpenPrice
                WS.Cells(Counter, 10).Value = YearlyChange
                
'               Formatting yearly change column after values are set
                If WS.Cells(Counter, 10).Value > 0 Then
                    WS.Cells(Counter, 10).Interior.ColorIndex = 4
                ElseIf WS.Cells(Counter, 10).Value < 0 Then
                    WS.Cells(Counter, 10).Interior.ColorIndex = 3
                End If
    
'               Calculating percent change - if statement to prevent dividing by 0 errors
                If OpenPrice <> 0 And Not IsNull(OpenPrice) Then
                    PercentChange = YearlyChange / OpenPrice
                Else
                    PercentChange = 0
                End If
    
'               Setting values for percent change and formatting to percent
                WS.Cells(Counter, 11).Value = PercentChange
                WS.Cells(Counter, 11).NumberFormat = "0.00%"
                
'               Setting values for total volume
                WS.Cells(Counter, 12).Value = Sum
                
'               Incrementing counter to move the next ticker calculations properly
                Counter = Counter + 1
    
'               Setting new ticker open price and reset volume count
                OpenPrice = WS.Cells(i + 1, 3).Value
                Sum = 0
            End If
        Next i
'       ------------Loop END for Ticker/YearlyChange/PercentChange------------

'       The above loop adds values to the same row that a change was found, so
'       the below will condence the table by removing extra spaces from columns I through K
        WS.Range("I:L").SpecialCells(xlCellTypeBlanks).Delete
       
'       Column Formatting
        ActiveSheet.UsedRange.EntireColumn.AutoFit
    Next WS
    
'   Calling Total Volume Calculation from another SUB
    'Call Volume

'   Calling Bonus from another SUB
    Call Bonus
    
    MsgBox ("Complete")
    
End Sub

Sub Bonus()

'   Setting variables
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim MaxVolume As Double
    
    For Each WS In Worksheets
        LastRow2 = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
'       Column Names
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
            
'       Setting max and min values
        MaxValue = Application.WorksheetFunction.Max(WS.Range("K:K"))
        MinValue = Application.WorksheetFunction.Min(WS.Range("K:K"))
        MaxVolume = Application.WorksheetFunction.Max(WS.Range("L:L"))
    
'       Setting values and formatting
        WS.Cells(2, 17).Value = MaxValue
        WS.Cells(3, 17).Value = MinValue
        WS.Cells(4, 17).Value = MaxVolume
        WS.Cells(2, 17).NumberFormat = "0.00%"
        WS.Cells(3, 17).NumberFormat = "0.00%"
        
'       ------------Loop BEGIN for Bonus----------
        For k = 2 To LastRow2
        
'           Check and set max ticker
            If WS.Cells(k, 11).Value = WS.Cells(2, 17).Value Then
                WS.Cells(2, 16).Value = WS.Cells(k, 9).Value
    
'           Check and set min ticker
            ElseIf WS.Cells(k, 11).Value = WS.Cells(3, 17).Value Then
                WS.Cells(3, 16).Value = WS.Cells(k, 9).Value
                
'           Check and set max volume ticker
            ElseIf WS.Cells(k, 12).Value = WS.Cells(4, 17).Value Then
                WS.Cells(4, 16).Value = WS.Cells(k, 9).Value
            End If
        Next k
'       ------------Loop END for Bonus------------
    
'   Column Formatting
    ActiveSheet.UsedRange.EntireColumn.AutoFit
    Next WS
End Sub


