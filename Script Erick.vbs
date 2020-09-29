  Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim I As Integer


         WS_Count = ActiveWorkbook.Worksheets.Count

         For I = 1 To WS_Count

            Call extract_tickers
            
            MsgBox ActiveWorkbook.Worksheets(I).Name

         Next I

      End Sub

Sub extract_tickers()

Dim a As String

b = Cells(Rows.Count, 1).End(xlUp).Row

Range("J1") = "Ticker"

Range("A2:A" & b).Copy Range("J2:J" & b)

Range("J2:J" & b).RemoveDuplicates (1)

Call data_extraction

End Sub
Sub data_extraction()

Large_list_limit = Cells(Rows.Count, 1).End(xlUp).Row
Short_list_limit = Cells(Rows.Count, 10).End(xlUp).Row

Dim volume As Double

'MsgBox (Str(Large_list_limit))
'MsgBox (Str(Short_list_limit))

Range("K1") = "Open Price Jan 1"
Range("L1") = "Close Price Dec 31"
Range("M1") = "Substraction"
Range("N1") = "Percent change"
Range("o1") = "Volume"

'Year extraction
u = Left(Cells(2, 2), 4)
Jan1 = Str(u) + "0101"
Dec31 = Str(u) + "1231"
Dec30 = Str(u) + "1230"

Cells(20, 100) = Jan1
Cells(21, 100) = Dec31
Cells(22, 100) = Dec30

'MsgBox (yeardate)

'MsgBox ("Year for this sheet: " & Str(u))

'Rows for short list

For x = 2 To Short_list_limit

volume = 0

'Rows for large list
    For y = 2 To Large_list_limit
    
    'MsgBox (Cells(x, 10))
    'MsgBox (Cells(y, 2))
    
        If Cells(x, 10).Value = Cells(y, 1).Value And Cells(20, 100).Value = Cells(y, 2).Value Then
        
            Cells(x, 11).Value = Cells(y, 3).Value
            
        End If
                
        If Cells(x, 10).Value = Cells(y, 1).Value Then
        
        volume = volume + Cells(y, 7)
        
        'MsgBox ("Current volume: " + Trim(Str(volume)))
        
        End If

        If Cells(x, 10).Value = Cells(y, 1).Value And (Cells(21, 100).Value = Cells(y, 2).Value Or Cells(22, 100).Value = Cells(y, 2).Value) Then
            
            Cells(x, 12).Value = Cells(y, 6).Value
        
        End If

    Next y

        'Resta
        Cells(x, 13) = (Cells(x, 12) - Cells(x, 11))
        
        'Cambio porcentual
        If (Cells(x, 11).Value = 0) Then
        
            Cells(x, 14) = 0
        
        Else
        
        Cells(x, 14) = ((Cells(x, 12) - Cells(x, 11)) / Cells(x, 11))

        End If

        If Cells(x, 14) > 0 Then
        
        Cells(x, 14).Interior.ColorIndex = 4
        
        Else
        
        Cells(x, 14).Interior.ColorIndex = 3
        
        End If
        
        Cells(x, 14).NumberFormat = "0.00%"
        
        Cells(x, 15) = volume

Next x

Call Min

Call Max

Call Vol


End Sub

Sub Min()

Short_list_limit = Cells(Rows.Count, 10).End(xlUp).Row

Cells(1, 19) = "Ticker"
Cells(1, 20) = "Value"
Cells(2, 18) = "Greatest decrease"

For t = 2 To Short_list_limit

    If Cells(t, 14).Value < Cells(t - 1, 14).Value Then
    
        Cells(2, 19) = Cells(t, 10)
        Cells(2, 20) = Cells(t, 14)
        
    End If

Next t

End Sub

Sub Max()

Short_list_limit = Cells(Rows.Count, 10).End(xlUp).Row

Cells(3, 18) = "Greatest increase"

For t = 2 To Short_list_limit

    If Cells(t, 14) > Cells(t - 1, 14) Then
    
        Cells(3, 19) = Cells(t, 10)
        Cells(3, 20) = Cells(t, 14)
        
    End If

Next t

End Sub

Sub Vol()

Short_list_limit = Cells(Rows.Count, 10).End(xlUp).Row

Cells(4, 18) = "Greatest volume"

For t = 2 To Short_list_limit

    If Cells(t, 15) > Cells(t + 1, 15) Then
    
        Cells(4, 19) = Cells(t, 10)
        Cells(4, 20) = Cells(t, 15)
        
    End If

Next t

End Sub
