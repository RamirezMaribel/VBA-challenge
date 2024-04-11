Attribute VB_Name = "Module1"
Sub ticker_mess()

' Loop through all worksheets
For Each ws In Worksheets
    With ws
    
    ' posting headers
    ticker_table = 2
    ws.Cells(1, "i").Value = "Ticker"
    ws.Cells(1, "j").Value = "Yearly Change"
    ws.Cells(1, "k").Value = "Percent Change"
    ws.Cells(1, "l").Value = "Total Stock Volume"
    ws.Cells(1, "m").Value = "Stored Open Price"
    
    ' Setting Boolean as described in Stackoverflow: https://stackoverflow.com/questions/59461571/how-do-i-keep-initial-value-in-a-for-loop
    Dim Open_Store As Boolean
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through all tickers
        For i = 2 To LastRow
            ' formatting precent stackoverflow: https://stackoverflow.com/questions/6854764/formatting-cells-as-percentage
            ws.Cells(i, "k").NumberFormat = "00.0%"
            ticker_ = ws.Cells(i, "a")
            Open_ = ws.Cells(i, "c").Value
            close_ = ws.Cells(i, "f")
    ' First if question stores the first opening value
         If Open_Store = False Then
         
         ' Posts opening value in column j
            ws.Range("m" & ticker_table).Value = Open_
            Open_Store = True
         End If
         
     ' Second if quesiton to bring in tickers in column i
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            Vol_ = Vol_ + ws.Cells(i, "g").Value
        
        ' Posts ticker to column i when ticker changes
           ws.Range("i" & ticker_table).Value = ticker_
      
        ' Posts stock volumne
            ws.Range("l" & ticker_table).Value = Vol_
                                  
        ' Resets the Volume total
             Vol_ = 0
             
        'Calculates percent change
            ws.Range("k" & ticker_table).Value = (close_ - ws.Range("m" & ticker_table).Value) / ws.Range("m" & ticker_table).Value
             
        'Calculates Yearly change
            ws.Range("j" & ticker_table).Value = close_ - ws.Range("m" & ticker_table).Value
        
        ' Resets the Open Store
            Open_Store = False
            
        ' Adds 1 to the ticker table to move the next ticker to the next line in column i
            ticker_table = ticker_table + 1
            
        Else
           
        ' Posts total to column l
            Vol_ = Vol_ + ws.Cells(i, "g").Value
        
        End If
        
       'Third & Fourth if question to format yearly change and precent change
        If ws.Cells(i, "j").Value < 0 Then
            ws.Cells(i, "j").Interior.ColorIndex = 3
        ElseIf ws.Cells(i, "j").Value > 0 Then
            ws.Cells(i, "j").Interior.ColorIndex = 4
        End If
        If ws.Cells(i, "k").Value < 0 Then
            ws.Cells(i, "k").Interior.ColorIndex = 3
        ElseIf ws.Cells(i, "k").Value > 0 Then
            ws.Cells(i, "k").Interior.ColorIndex = 4
        End If
     Next i
    End With
 
    Next ws

End Sub

