Sub new1()
Dim WS As Worksheet

Dim symbol As String

Dim volume As Long
volume = 0

Dim rownum As Long

Dim lastnumrow As Long

Dim opening As Double

Dim closing As Double

Dim mainrow As Integer

Dim newsymbol As Integer

Dim stockvolume As LongLong

Dim percentchange As Double



For Each WS In Worksheets
WS.Activate
newsymbol = 0
opening = 0
closing = 0
stockvolume = 0
mainrow = 2

Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"
 

numlastrow = Cells(Rows.Count, 1).End(xlUp).Row
For rownum = 2 To numlastrow
    If newsymbol = 0 Then
    
        opening = Cells(rownum, 3).Value
        newsymbol = 1
    End If
    
    stockvolume = stockvolume + Cells(rownum, 7).Value
    
    If Cells(rownum + 1, 1).Value <> Cells(rownum, 1).Value Then
        newsymbol = 0
    
        symbol = Cells(rownum, 1).Value
        Range("I" & mainrow).Value = symbol
        
        closing = Cells(rownum, 6).Value
        Range("J" & mainrow).Value = (closing - opening)
        
        If opening <> 0 Then
            percentchange = (closing - opening) / opening
        Else
            percentchange = 0
        End If
        
        Range("K" & mainrow).Value = FormatPercent(percentchange)
        
        Range("L" & mainrow).Value = stockvolume
        
        mainrow = mainrow + 1
        stockvolume = 0
    
    End If
    
    If Cells(rownum, 10).Value < 0 Then
        Cells(rownum, 10).Interior.ColorIndex = 3
    Else
        Cells(rownum, 10).Interior.ColorIndex = 4
    End If
    
    
Next rownum

Next WS

End Sub
