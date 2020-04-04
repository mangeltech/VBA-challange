Attribute VB_Name = "Module1"

Sub Ticker()

'define variables
Dim ws As Worksheet
Dim Ticker As String
Dim vol As Double
vol = 0
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim Summary_Table_Row As Integer

'this preventing overflow error
On Error Resume Next

'run through each worksheet
For Each ws In ThisWorkbook.Worksheets
    'setting headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
' SettingAdditional Headers
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
    
    'setting up integers for loop
    Summary_Table_Row = 2

    'loop
        For i = 2 To ws.UsedRange.Rows.Count
            If year_open = 0 Then

                    year_open = ws.Cells(i, 3).Value
            End If

        If ws.Cells(i, 1) = ws.Cells(i, 1) And ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               'finding all the values
             year_close = ws.Cells(i, 6).Value
             yearly_change = year_close - year_open
               percent_change = (yearly_change) / year_open
             year_open = ws.Cells(i + 1, 3).Value
             
          
            Ticker = ws.Cells(i, 1).Value
            vol = vol + ws.Cells(i, 7).Value

            
            

            'inserting values
            ws.Range("I" & Summary_Table_Row).Value = Ticker
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Range("L" & Summary_Table_Row).Value = vol
            Summary_Table_Row = Summary_Table_Row + 1

             vol = 0
            Else
                vol = vol + ws.Cells(i, 7).Value
        
            End If
'formating color column
            If ws.Cells(i, 10) < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 10
            End If
'finishing loop
    Next i
    
    'formating yearly percent
    ws.Columns("K").NumberFormat = "0.00%"


    


'moving to next worksheet
Next ws


End Sub


