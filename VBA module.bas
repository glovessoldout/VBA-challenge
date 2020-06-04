Attribute VB_Name = "Module1"
Sub ticker()
    Dim lasttick As Long
    lasttick = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (lasttick)
    
    Dim changecount As Long
    changecount = 0
    
    'My version of excel doesn't support LongLong type, probably because it's version 2010. From googling, the LongLong type should be the solution that I can't use.
    'Dim stockvolume as LongLong
    Dim stockvolume As Long
    stockvolume = 0
    
    Dim alpha, omega As Long
    alpha = 0
    omega = 0
    
    Cells(1, 9).Value = "Ticker Symbol"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    
    tickcount = 2
    For i = 2 To lasttick
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'ticker symbol collection
            Cells(tickcount, 9).Value = Cells(i, 1).Value
            'yearly change collection
            Cells(tickcount, 10).Value = Cells(i, 6).Value - Cells(i - changecount, 3).Value
            ' conditional formatting
            If Cells(tickcount, 10).Value >= 0 Then
                Cells(tickcount, 10).Interior.ColorIndex = 4
                Else
                Cells(tickcount, 10).Interior.ColorIndex = 3
            End If
            ' percent change
            Cells(tickcount, 11).Value = Cells(i - changecount, 3).Value / Cells(i, 6).Value
            Cells(tickcount, 11).Value = FormatPercent(Cells(tickcount, 11).Value, 2)
            ' total stock volume, doesn't work on my version of excel because it's 32bit
           ' For j = (i - changecount) To i
           '     stockvolume = stockvolume + Cells(j, 7)
           ' Next
            
            'stockvolume =
            'stockvolume Cells(i, 7).Value
            'Cells(tickcount, 12).Value = stockvolume
            
            
            'counting indexes
            tickcount = tickcount + 1
            changecount = 0
            stockvolume = 0
        Else
            changecount = changecount + 1
            'stockvolume = stockvolume + Cells(i, 7).Value
        End If
    Next
    
End Sub
