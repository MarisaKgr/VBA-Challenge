Attribute VB_Name = "Module1"
Sub stock_data()

Dim ws As Worksheet
Dim lastrow As Long

Worksheets("2016").Columns("I:L").Clear
Worksheets("2015").Columns("I:L").Clear
Worksheets("2016").Columns("I:L").Clear



'add column names
For Each ws In Worksheets

    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Columns("I:Q").AutoFit
    
    Dim Ticker As String
    Dim Total As Double
    Dim summary As Long
    Dim open_p As Double
    Dim end_p As Double
    Dim change As Double
    Dim percent As Double
    Dim x As Long
    
    'populate ticker and total_vol in summary starting position 2
    summary = 2

    'opening price for each ticker
    open_p = ws.Cells(2, 3).Value

    For x = 2 To lastrow
 
        'pull total from column 7
        Total = Total + ws.Cells(x + 1, 7).Value
              
        ' generate ticker from column 1 to last row
        If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
            
            Ticker = ws.Cells(x, 1).Value
                           
            'calculate yearly change
            end_p = ws.Cells(x, 6).Value


            change = end_p - open_p
    
            'reset total
            Total = 0
            
            'set as zero if open and close are zero to avoid calculation error
            If (open_p = 0 And end_p = 0) Then
            percent = 0
            'calculate percent change
            ElseIf (open_p <> 0 And end_p <> 0) Then
            percent = (change / open_p)
    
        
        
            'set positive change to green
            If change > 0 Then
                ws.Range("J" & summary).Interior.ColorIndex = 4
                'set negative or no change to red
                ElseIf change <= 0 Then
                ws.Range("J" & summary).Interior.ColorIndex = 3
            End If
    
    
            End If
 
        'add to summary to move to next row
        summary = summary + 1
         
        'move to next ticker opening price
        open_p = ws.Cells(x + 1, 3)
    
        'report values back to next position in summary table
        ws.Range("I" & summary - 1).Value = Ticker
        ws.Range("J" & summary - 1).Value = change
        ws.Range("K" & summary - 1).Value = percent
        'format cells to percent
        ws.Range("K" & summary - 1).NumberFormat = "0.00%"
    
    End If
        
        'populate values into summary table post calulation loop
        ws.Range("L" & summary).Value = Total
    
    Next x



Next ws

'Bonus
Worksheets("2016").Columns("O:Q").Clear
Worksheets("2015").Columns("O:Q").Clear
Worksheets("2016").Columns("O:Q").Clear

For Each ws In Worksheets

ws.Range("O2") = "Greatest % Increase"
ws.Range("O3") = "Greatest % Decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

ws.Columns("I:Q").AutoFit



'find greatest decrease
Set ws = Worksheets("2016")

ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K10000"), 1)

ws.Range("Q3").NumberFormat = "0.00%"


Set ws = Worksheets("2014")

ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K10000"), 1)

ws.Range("Q3").NumberFormat = "0.00%"

Set ws = Worksheets("2015")

ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K2:K10000"), 1)

ws.Range("Q3").NumberFormat = "0.00%"


'return highest %

Set ws = Worksheets("2016")

ws.Range("Q2") = Application.WorksheetFunction.Large(ws.Range("K2:K10000"), 1)

ws.Range("Q2").NumberFormat = "0.00%"

Set ws = Worksheets("2014")

ws.Range("Q2") = Application.WorksheetFunction.Large(ws.Range("K2:K10000"), 1)

ws.Range("Q2").NumberFormat = "0.00%"

Set ws = Worksheets("2015")

ws.Range("Q2") = Application.WorksheetFunction.Large(ws.Range("K2:K10000"), 1)

ws.Range("Q2").NumberFormat = "0.00%"

'return highest total

Set ws = Worksheets("2016")

ws.Range("Q4") = Application.WorksheetFunction.Large(ws.Range("l2:l10000"), 1)


Set ws = Worksheets("2014")

ws.Range("Q4") = Application.WorksheetFunction.Large(ws.Range("l2:l10000"), 1)


Set ws = Worksheets("2015")

ws.Range("Q4") = Application.WorksheetFunction.Large(ws.Range("l2:l10000"), 1)
  

Next ws
'Matching tickers to calculated values. t = 2014, t2 = 2015, t3 = 2016

For Each ws In Worksheets

Dim t As Integer
Dim t2 As Integer
Dim t3 As Integer

For t = 2 To ActiveWorkbook.Worksheets("2014").Range("K1").CurrentRegion.Rows.Count
            If ActiveWorkbook.Worksheets("2014").Cells(t, 12).Value = ActiveWorkbook.Worksheets("2014").Cells(4, 17).Value Then
                ActiveWorkbook.Worksheets("2014").Cells(4, 16).Value = ActiveWorkbook.Worksheets("2014").Cells(t, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2014").Cells(t, 11).Value = ActiveWorkbook.Worksheets("2014").Cells(3, 17).Value Then
                ActiveWorkbook.Worksheets("2014").Cells(3, 16).Value = ActiveWorkbook.Worksheets("2014").Cells(t, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2014").Cells(t, 11).Value = ActiveWorkbook.Worksheets("2014").Cells(2, 17).Value Then
                ActiveWorkbook.Worksheets("2014").Cells(2, 16).Value = ActiveWorkbook.Worksheets("2014").Cells(t, 9).Value
            End If

Next t

For t2 = 2 To ActiveWorkbook.Worksheets("2015").Range("K1").CurrentRegion.Rows.Count
            If ActiveWorkbook.Worksheets("2015").Cells(t2, 12).Value = ActiveWorkbook.Worksheets("2015").Cells(4, 17).Value Then
                ActiveWorkbook.Worksheets("2015").Cells(4, 16).Value = ActiveWorkbook.Worksheets("2015").Cells(t2, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2015").Cells(t2, 11).Value = ActiveWorkbook.Worksheets("2015").Cells(3, 17).Value Then
                ActiveWorkbook.Worksheets("2015").Cells(3, 16).Value = ActiveWorkbook.Worksheets("2015").Cells(t2, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2015").Cells(t2, 11).Value = ActiveWorkbook.Worksheets("2015").Cells(2, 17).Value Then
                ActiveWorkbook.Worksheets("2015").Cells(2, 16).Value = ActiveWorkbook.Worksheets("2015").Cells(t2, 9).Value
            End If

Next t2
            
For t3 = 2 To ActiveWorkbook.Worksheets("2016").Range("K1").CurrentRegion.Rows.Count
            If ActiveWorkbook.Worksheets("2016").Cells(t3, 12).Value = ActiveWorkbook.Worksheets("2016").Cells(4, 17).Value Then
                ActiveWorkbook.Worksheets("2016").Cells(4, 16).Value = ActiveWorkbook.Worksheets("2016").Cells(t3, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2016").Cells(t3, 11).Value = ActiveWorkbook.Worksheets("2016").Cells(3, 17).Value Then
                ActiveWorkbook.Worksheets("2016").Cells(3, 16).Value = ActiveWorkbook.Worksheets("2016").Cells(t3, 9).Value
            ElseIf ActiveWorkbook.Worksheets("2016").Cells(t3, 11).Value = ActiveWorkbook.Worksheets("2016").Cells(2, 17).Value Then
                ActiveWorkbook.Worksheets("2016").Cells(2, 16).Value = ActiveWorkbook.Worksheets("2016").Cells(t3, 9).Value
            End If

Next t3

Next ws

End Sub

