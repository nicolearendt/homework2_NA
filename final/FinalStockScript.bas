Attribute VB_Name = "Module1"
Sub FinalStockScript()
    
    For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
  
        'Find Last Row and declare variables
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Dim FirstPrice As Double
        Dim LastPrice As Double
        Dim FirstRowNumber As Double
        Dim LastRowNumber As Double
        Dim TickName As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        
        'Put in Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        'Declare and Initialize Column number
        Dim num As Integer
        num = 2
  
        ' Go through each row
        For i = 1 To LastRow
            
            'Check if you are at the beginning of a new worksheet
            'and set the open price, the row number of the first instance, and the ticker name
            If Cells(i, 1).Value Like "<*" Then
                
                FirstPrice = Cells(i + 1, 3).Value
                FirstRowNumber = i + 1
                TickName = Cells(i + 1, 1).Value
            
            'Find the last value of the stock
            ElseIf (Cells(i, 1).Value <> Cells(i + 1, 1).Value) Then
                
                'Put new ticker name in our list
                Cells(num, 9).Value = TickName
                
                'Set the last price
                LastPrice = Cells(i, 6).Value
                
                'Calculate the yearly change and determine if
                'its a positve or negative change
                YearlyChange = LastPrice - FirstPrice
                Cells(num, 10).Value = YearlyChange
                If (Cells(num, 10).Value > 0 Or Cells(num, 10).Value = 0) Then
                    Cells(num, 10).Interior.ColorIndex = 4
                Else
                    Cells(num, 10).Interior.ColorIndex = 3
                End If
                
                'Calculate percent change and format the cell
                If (FirstPrice = 0) Then
                    PercentChange = 0
                    Cells(num, 11).Value = PercentChange
                    Cells(num, 11).NumberFormat = "0.00%"
                Else
                    PercentChange = YearlyChange / FirstPrice
                    Cells(num, 11).Value = PercentChange
                    Cells(num, 11).NumberFormat = "0.00%"
                End If
                
                'Set the new Ticker name and new Open Price
                TickName = Cells(i + 1, 1).Value
                FirstPrice = Cells(i + 1, 3).Value
                
                'Set number of row of the last instance and Calculate the Total Volume
                LastRowNumber = i
                Cells(num, 12).Formula = "=SUM(" & Range(Cells(FirstRowNumber, 7), Cells(LastRowNumber, 7)).Address(False, False) & ")"
                
                'Set the new first row number and new row number in our list
                FirstRowNumber = i + 1
                num = num + 1

            End If
  
        Next i
        
        'Find the Last row in Column J
        JLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'Declare Variables to find the greatest precent Inc./Dec.
        'as well as greatest total stock volume
       Dim Big As Double
       Dim BigTicker As String
       Dim Small As Double
       Dim SmallTicker As String
       Dim BigTotal As Double
       Dim BigTotalTicker As String
        
        'Initialize variables
        Big = 0
        Small = 100000
        BigTotal = 0
        
        'Go through Each Row
        For j = 2 To JLastRow
        
            'Find the greatest % increase
            If (Cells(j, 11).Value > Big) Then
                Big = Cells(j, 11).Value
                BigTicker = Cells(j, 9).Value
            End If
            
            'Find the greatest % decrease
            If (Cells(j, 11).Value < Small) Then
                Small = Cells(j, 11).Value
                SmallTicker = Cells(j, 9).Value
            End If
        
            'Find the greatest total volume
            If (Cells(j, 12).Value > BigTotal) Then
                BigTotal = Cells(j, 12).Value
                BigTotalTicker = Cells(j, 9).Value
            End If
        
        Next j
        
        'Update our graph
        Cells(2, 16).Value = BigTicker
        Cells(2, 17).Value = Big
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 16).Value = SmallTicker
        Cells(3, 17).Value = Small
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 16).Value = BigTotalTicker
        Cells(4, 17).Value = BigTotal
  
    Next ws
  
End Sub
