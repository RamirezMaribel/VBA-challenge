Attribute VB_Name = "Module2"
Sub variableinfo()
For Each ws In Worksheets
    With ws
     

' posting headers for the variable info table
    ws.Cells(1, "p").Value = "Ticker"
    ws.Cells(1, "q").Value = "Value"
    ws.Cells(2, "o").Value = "Greatest % Increase"
    ws.Cells(3, "o").Value = "Greatest % Decrease"
    ws.Cells(4, "o").Value = "Greatest Total Volume"


    
'counting last rows of 1st summary table

        LastRow2 = ws.Cells(Rows.Count, "k").End(xlUp).Row
        For i = 2 To LastRow2
        
        Set Range1 = ws.Range("k2:k" & LastRow2)
        Set Range2 = ws.Range("l2:l" & LastRow2)
        ws.Cells(2, "q").NumberFormat = "00.0%"
        ws.Cells(3, "q").NumberFormat = "00.0%"
    ' reference: https://stackoverflow.com/questions/42633273/finding-max-of-a-column-in-vba
        Max = Application.WorksheetFunction.Max(Range1)
        Min = Application.WorksheetFunction.Min(Range1)
        greatest = Application.WorksheetFunction.Max(Range2)
        
        
        ' if that finds the max
        If ws.Cells(i, "k").Value = Max Then

        ' posting to variable info table
            ws.Cells(2, "p").Value = ws.Cells(i, "i")
            ws.Cells(2, "q").Value = Max

        End If

         ' if that finds the min
         If ws.Cells(i, "k").Value = Min Then
    
        ' posting to variable info table
            ws.Cells(3, "p").Value = ws.Cells(i, "i")
            ws.Cells(3, "q").Value = Min
        End If

        ' if that finds t he greatest value
        If ws.Cells(i, "l").Value = greatest Then
        
        ' Posting to variable info table
            ws.Cells(4, "p").Value = ws.Cells(i, "i")
            ws.Cells(4, "q").Value = greatest
        
        End If

        Next i
    End With
 
    Next ws

End Sub





