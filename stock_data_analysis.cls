VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Author: Esubalew Adal
'This VBA code updates all the worksheets in this workbook

Sub VBA_Challenge()

'Variable diclaration
Dim ws As Worksheet
'Dim VBAProject As Workbook
Dim Summary_Table_Row As Integer
Dim ClosePrice As Double
Dim OpenPrice As Double
Dim Percent_Change As Double
Dim Yearly_Change As Double
Dim Ticker_Total_Volume As Double
Dim Greatest_Percent_Increase As Double
Dim Greatest_Percent_Decrease As Double
Dim Greatest_Percent_Decrease_Ticker As String
Dim Geatest_Volume_Ticker As String
Dim Greatest_Total_Volume As Double
Dim TickerName As String
Dim Greatest_Percent_Increase_Ticker As String

For Each ws In Worksheets
'Initializing variables
MsgBox ("Now updating:" + ws.Name)
Greatest_Percent_Increase = 0
Greatest_Percent_Decrease = 0
Greatest_Total_Volume = 0
Ticker_Total_Volume = 0
Summary_Table_Row = 2
Greatest_Percent_Increase_Ticker = ""
Greatest_Percent_Decrease_Ticker = ""
Geatest_Volume_Ticker = ""
TickerName = ""

' Get the last row
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Create Summary Table Headers

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("I1:Q1").Columns.AutoFit
ws.Range("O2:O4").Columns.AutoFit
ws.Range("A1:Q1").Font.Bold = True

OpenPrice = ws.Cells(2, 3).Value   'For capturing the first ticker (AAB)open price just for this mehtod only.
                                   'If using the array method,or other methods, explicitly assigning the openPrice value
                                   'is not neccessary


'Iteration for the current worksheet rows

For i = 2 To lastrow

' Comparing the adjacent rows values

If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then

'If this condition is true, do this

TickerName = ws.Cells(i, 1).Value

ClosePrice = ws.Cells(i, 6).Value

Yearly_Change = ClosePrice - OpenPrice

Percent_Change = ((ClosePrice / OpenPrice) - 1) * 100

Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

'Populate the Summary Table Values

ws.Range("I" & Summary_Table_Row).Value = TickerName

ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

ws.Range("K" & Summary_Table_Row).Value = (Percent_Change) & "%"

ws.Range("L" & Summary_Table_Row).Value = Ticker_Total_Volume

'Adding color to the cells to indicate good (green)and bad (red)performing stocks

If (Yearly_Change > 0) Then

ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4

Else

ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3

End If
'Increment the Summary Table Rows

Summary_Table_Row = Summary_Table_Row + 1

'Reset variables to capture the next ticker OpenPrice and ClosePrice where it happens @ (i+1)

Yearly_Change = 0            'Resets the value every loop to capture the next ticker value
ClosePrice = 0               'Resets the value every loop to capture the next ticker value
OpenPrice = ws.Cells(i + 1, 3)  'capture the next ticker openPrice

'Comparison to determine the maximum and minimum percent changes

If Percent_Change > Greatest_Percent_Increase Then

Greatest_Percent_Increase = Percent_Change

Greatest_Percent_Increase_Ticker = TickerName

ElseIf Percent_Change < Greatest_Percent_Decrease Then

Greatest_Percent_Decrease = Percent_Change

Greatest_Percent_Decrease_Ticker = TickerName

End If
 'Determine the maximum volume
 
If Ticker_Total_Volume > Greatest_Total_Volume Then

Greatest_Total_Volume = Ticker_Total_Volume

Geatest_Volume_Ticker = TickerName

End If

Percent_Change = 0
Ticker_Total_Volume = 0

Else
'Otherwise, if the above condition is not True, do this repeatdely untilt the condition becomes true. The volumes values will be added row-by-row

Ticker_Total_Volume = Ticker_Total_Volume + ws.Cells(i, 7).Value

End If

'Populating the Summary_Table after the iteration

ws.Range("P2").Value = Greatest_Percent_Increase_Ticker
ws.Range("P3").Value = Greatest_Percent_Decrease_Ticker
ws.Range("P4").Value = Geatest_Volume_Ticker

ws.Range("Q2").Value = Greatest_Percent_Increase & "%"
ws.Range("Q3").Value = Greatest_Percent_Decrease & "%"
ws.Range("Q4").Value = Greatest_Total_Volume
Next i
Next ws
End Sub


