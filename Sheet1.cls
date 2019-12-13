VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Market()

Dim ws As Worksheet

For Each ws In Worksheets

Dim LastRow As Double
LastRow = ws.Range("G" & Rows.Count).End(xlUp).Row
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Stock_Volume As Double
Stock_Volume = 0
Dim Summary_Table As Integer
Summary_Table = 2
Dim Stock_Open As Double
Dim Stock_Close As Double
Year_Open = 0
Year_End = 0


For i = 2 To LastRow

    If Year_Open = 0 Then
    Year_Open = ws.Cells(i, 3).Value
    
    End If
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
         
        Ticker = ws.Cells(i, 1).Value
        Year_End = ws.Cells(i, 6).Value
        Yearly_Change = Year_End - Year_Open
        Stock_Volume = Stock_Volume + ws.Cells(i, 7)
        
        If Year_Open <> 0 Then
        
        Percent_Change = (Yearly_Change / Year_Open) * 100
        
        End If
        
        ws.Range("J" & Summary_Table).Value = Yearly_Change
        ws.Range("I" & Summary_Table).Value = Ticker
        ws.Range("K" & Summary_Table).Value = ((Percent_Change) & "%")
        ws.Range("L" & Summary_Table).Value = Stock_Volume
        
        If Yearly_Change > 0 Then
        ws.Range("J" & Summary_Table).Interior.ColorIndex = 4
        
        ElseIf Yearly_Change < 0 Then
        ws.Range("J" & Summary_Table).Interior.ColorIndex = 3
        
        End If
        
        Summary_Table = Summary_Table + 1
        Year_Open = 0
        Yearly_Change = 0
        Percent_Change = 0
        Stock_Volume = 0
        
        Else
         
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
   
    End If

    
Next i

Next ws

End Sub



