Attribute VB_Name = "Module3"
Sub Stock_Practice()

Dim ticker As String
Dim yearly_change As Double
    yearly_change = 0
Dim total_volume As Double
   total_volume = 0
Dim Summary_Table_Row As Integer
Dim open_number As Double
Dim close_number As Double
Dim percent_change As Double
Dim cell As Range
Dim ws As Worksheet




For Each ws In Worksheets
    ws.Activate
    Summary_Table_Row = 2
    
 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For j = 2 To lastrow

        If ws.Cells(j + 1, 1).Value <> ws.Cells(j, 1).Value Then
            ticker = ws.Cells(j, 1).Value
            total_volume = total_volume + ws.Cells(j, 7).Value
    
            open_number = open_number + ws.Cells(j, 3).Value
            close_number = close_number + ws.Cells(j, 6).Value
    
            yearly_change = yearly_change + (close_number - open_number)
   
            percent_change = percent_change + (yearly_change / open_number)
    
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Range("L" & Summary_Table_Row).Value = total_volume
    
            Summary_Table_Row = Summary_Table_Row + 1
    
                total_volume = 0
                yearly_change = 0
                percent_change = 0
            
                open_number = ws.Cells(j, 3).Value
                close_number = ws.Cells(j, 6).Value
                
    
        Else
             total_volume = total_volume + ws.Cells(j, 7).Value
             
            
            
        End If

Next j

For j = 2 To lastrow


    If ws.Cells(j, 10).Value > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    Else
        ws.Cells(j, 10).Interior.ColorIndex = 3
    End If


Next j

  For j = 2 To lastrow
    ws.Cells(j, "K").NumberFormat = "0.00%"
  
  Next j
  
Cells(1, "I").Value = "ticker"
Cells(1, "J").Value = "yearly change"
Cells(1, "K").Value = "Percent Change"
Cells(1, "L").Value = "Total Stock Volume"


Next ws
End Sub

