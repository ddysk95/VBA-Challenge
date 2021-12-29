VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vba_hw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub vba_hw()

    ' loop thru sheets
    For Each ws In Worksheets


'Title new chart
ws.Range("J1").Value = "Ticker"
ws.Range("k1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Vol"

'ticker variable
Dim Tic As String

'ticker total
Dim TicTot As Double
TicTot = 0

'chart row
Dim r As Integer
r = 2

'last row
lastr = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'opening of that year
Dim opening As Double
    
'closing of that year
Dim closing As Double


'rows of ticker
Dim TickRows As Double
TickRows = 0

'for loop for worksheet
For i = 2 To lastr

    'check if ticker is same, if not:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    'Ticker
    Tic = ws.Cells(i, 1).Value
    
    'Ticker total
    TicTot = TicTot + ws.Cells(i, 7).Value
    
    'print to chart row
    ws.Range("J" & r).Value = Tic
    
    opening = ws.Cells(i - TickRows, 3).Value

    closing = ws.Cells(i, 6).Value

    ' print Yearly Change to chart
    ws.Range("k" & r).Value = closing - opening
    
    If opening <> 0 Then
    ' print Percentage Change to chart
    ws.Range("l" & r).Value = (closing - opening) / opening
    
    Else
    ws.Range("l" & r).Value = 0
    
    End If
    
       ' print total to chart
    ws.Range("m" & r).Value = TicTot
    
    'add one line to chart
    r = r + 1
    
    'reset total
    TicTot = 0
    
    'reset ticker rows
    TickRows = 0

'if same ticker
Else
'add to total
    TicTot = TicTot + ws.Cells(i, 7).Value
    TickRows = TickRows + 1

    
    End If
    
    
    Next i
'--------------------------------------------------------------------
    
'formatting/conditionals

For i = 2 To r

ws.Range("L" & i).Style = "Percent"

If ws.Range("K" & i).Value < 0 Then
ws.Range("K" & i).Interior.ColorIndex = 3
Else

ws.Range("K" & i).Interior.ColorIndex = 4
End If

' titles for 2nd chart
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = " Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("Q3").Style = "Percent"
ws.Range("Q2").Style = "Percent"

'Greatest percentage
If ws.Range("L" & i).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Range("L" & i).Value
    ws.Range("P3").Value = ws.Range("J" & i).Value
    
    ElseIf ws.Range("L" & i).Value > ws.Range("Q2").Value Then
    ws.Range("Q2").Value = ws.Range("L" & i).Value
    ws.Range("P2").Value = ws.Range("J" & i).Value
End If

'Greatest total
If ws.Range("M" & i).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Range("m" & i).Value
    ws.Range("P4").Value = ws.Range("J" & i).Value
End If
    

Next i

Next ws

End Sub
