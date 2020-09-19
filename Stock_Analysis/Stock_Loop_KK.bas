Attribute VB_Name = "Stock_Loop_KK"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Ticker_Loop
    Next
    Application.ScreenUpdating = True
End Sub



Sub Ticker_Loop()

Dim alltickers As String
Dim uniquetickers As Integer
uniquetickers = 2

Range("I1").Value = "Tickers"

For I = 2 To 797711

    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
        alltickers = Cells(I, 1).Value
        
        Range("I" & uniquetickers).Value = alltickers
        
        uniquetickers = uniquetickers + 1
        
    End If
Next I
    



Dim opening As Double
Dim closing As Double
Dim yearly_change As Double
yearly_change = 2

Range("J1").Value = "Yearly Change"

For I = 2 To 797711

    If Cells(I + 1, 2).Value > Cells(I, 2).Value And Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    
        opening = Cells(I, 3).Value
        
    ElseIf Cells(I - 1, 2).Value < Cells(I, 2).Value And Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
        closing = Cells(I, 6).Value
            
        Range("J" & yearly_change).Value = (closing - opening)
        
        yearly_change = yearly_change + 1
        
    End If
    
Next I



Dim percent_change_row As Double
percent_change_row = 2

Range("K1").Value = "Percent Change"

For I = 2 To 797711

           
    If Cells(I + 1, 2).Value > Cells(I, 2).Value And Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    
        opening = Cells(I, 3).Value
        
    ElseIf Cells(I - 1, 2).Value < Cells(I, 2).Value And Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        
        closing = Cells(I, 6).Value
            
        If opening = 0 Then
        
           Range("K" & percent_change_row).Value = 0
           
           percent_change_row = percent_change_row + 1
           
           Else
           
            Range("K" & percent_change_row).Value = ((closing - opening) / opening)
        
            percent_change_row = percent_change_row + 1
        
        End If
        
    End If
    
Next I


Dim total_volume As Double
total_volume = 0

Dim total_volume_row As Integer
total_volume_row = 2

Range("L1").Value = "Total Volume"

For I = 2 To 797711

    If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
    
            
        total_volume = total_volume + Cells(I, 7).Value
        
        Range("L" & total_volume_row).Value = total_volume
        
        total_volume_row = total_volume_row + 1
        
        total_volume = 0
        
    Else
    
        total_volume = total_volume + Cells(I, 7).Value
        
    End If
    
Next I
        
        


Range("N2").Value = "Greatest % Increase"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

max_increase = Application.WorksheetFunction.Max(Range("K:K"))

Range("P2").Value = max_increase

For I = 2 To 797711

    If Cells(I, 11).Value = max_increase Then
    
        Ticker = Cells(I, 9).Value
        
        Range("O2").Value = Ticker
    
    End If
    
Next I




Range("N3").Value = "Greatest % Decrease"


max_decrease = Application.WorksheetFunction.Min(Range("K:K"))

Range("P3").Value = max_decrease

For I = 2 To 797711

    If Cells(I, 11).Value = max_decrease Then
    
        Ticker = Cells(I, 9).Value
        
        Range("O3").Value = Ticker
    
    End If
    
Next I


Range("N4").Value = "Greatest Total Volume"


greatest_total_volume = Application.WorksheetFunction.Max(Range("L:L"))

Range("P4").Value = greatest_total_volume

For I = 2 To 797711

    If Cells(I, 12).Value = greatest_total_volume Then
    
        Ticker = Cells(I, 9).Value
        
        Range("O4").Value = Ticker
    
    End If
    
Next I
End Sub

