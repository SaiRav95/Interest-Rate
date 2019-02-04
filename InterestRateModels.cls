VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Function NRoot(x As Double, y As Integer) As Double
    NRoot = x ^ (1 / y)
End Function



Sub Vasicek()

Dim x As Integer

x = InputBox("Enter a number for the Interest model")

If x = 1 Then

'Vasicek Model'

    Dim a As Double
    Dim b As Double
    Dim Sigma As Double
    Dim r As Double
    Dim n As Integer
    Dim dt As Double
    Dim dr As Double
    Dim i As Integer
    Dim t As Integer


    a = InputBox("Enter a value for a")
    b = InputBox("Enter a value for b")
    Sigma = InputBox("Enter a value for Sigma")
    r = InputBox("Enter a value for r")
    t = InputBox("Enter a value for t")
    n = InputBox("Enter a value for n")


    dt = t / n

    For i = 1 To n

        dr = a * (b - r) * dt + Sigma * Excel.WorksheetFunction.NormInv(Rnd(), 0, 1) * NRoot(dt, 2)
        r = r + dr
    
        Cells(i, 2).Value = r
    
    Next i

    For i = 1 To n

        Cells(i, 1).Value = i

    Next i

ElseIf x = 2 Then

'CIR Model'


    a = InputBox("Enter a value for a")
    b = InputBox("Enter a value for b")
    Sigma = InputBox("Enter a value for Sigma")
    r = InputBox("Enter a value for r")
    t = InputBox("Enter a value for t")
    n = InputBox("Enter a value for n")


    dt = t / n

    For i = 1 To n

        dr = a * (b - r) * dt + Sigma * NRoot(r, 2) * Excel.WorksheetFunction.NormInv(Rnd(), 0, 1) * NRoot(dt, 2)
        r = r + dr
    
        Cells(i, 2).Value = r
    
    Next i

    For i = 1 To n

        Cells(i, 1).Value = i

    Next i
    
    
ElseIf x = 3 Then

    'Rendleman and Barter Model'
    
    b = InputBox("Enter a value for b")
    Sigma = InputBox("Enter a value for Sigma")
    r = InputBox("Enter a value for r")
    t = InputBox("Enter a value for t")
    n = InputBox("Enter a value for n")
    
    dt = t / n

    For i = 1 To n

        dr = b * r * dt + Sigma * r * Excel.WorksheetFunction.NormInv(Rnd(), 0, 1) * NRoot(dt, 2)
        r = r + dr
    
        Cells(i, 2).Value = r
    
    Next i

    For i = 1 To n

        Cells(i, 1).Value = i

    Next i
    
    
Else

    MsgBox ("Invalid Input")
    
End If
    
    

End Sub

