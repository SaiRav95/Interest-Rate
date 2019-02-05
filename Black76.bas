Attribute VB_Name = "Module1"
Option Explicit

Function NRoot(x As Double, y As Integer) As Double
    NRoot = x ^ (1 / y)
End Function

Sub Black76()

Dim Loan As Double
Dim Cap As Double
Dim Floor As Double
Dim Caplet As Double
Dim ForwardRate As Double
Dim Sigma As Double
Dim InterestRate As Double
Dim Time1 As Double
Dim Time2 As Double
Dim d1 As Double
Dim d2 As Double
Dim x As Integer

'x = 1 for Interest Rate Caps'
'x = 2 for Interest Rate Floor'


x = InputBox("Enter the number associated to the Interest Rate you wanna call")

If x = 1 Then

    Loan = InputBox("Enter the Loan Amount")
    Cap = InputBox("Enter the Cap rate which is set")
    ForwardRate = InputBox("Enter the rate at Future")
    Sigma = InputBox("Enter the Volatility")
    InterestRate = InputBox("Enter the interest rate (Continuous)")
    Time1 = InputBox("Enter the Time when payoff is calculated")
    Time2 = InputBox("Enter the Time when the payoff actually occured")

    d1 = (Log(ForwardRate / Cap) + (Sigma ^ 2) * Time1 / 2) / Sigma * NRoot(Time1, 2)
    d2 = d1 - Sigma * NRoot(Time1, 2)

    Caplet = Loan * (Time2 - Time1) * (Exp(-InterestRate * Time2)) * ((ForwardRate * Excel.WorksheetFunction.NormDist(d1, 0, 1, True)) - (Cap * Excel.WorksheetFunction.NormDist(d2, 0, 1, True)))

    If Caplet > 0 Then

    MsgBox ("The value of the Caplet" & " " & Caplet)

    ElseIf Caplet <= 0 Then

    MsgBox ("The value of the Caplet" & " " & 0)

ElseIf x = 2 Then

    Loan = InputBox("Enter the Loan Amount")
    Floor = InputBox("Enter the Floor rate which is set")
    ForwardRate = InputBox("Enter the rate at Future")
    Sigma = InputBox("Enter the Volatility")
    InterestRate = InputBox("Enter the interest rate (Continuous)")
    Time1 = InputBox("Enter the Time when payoff is calculated")
    Time2 = InputBox("Enter the Time when the payoff actually occured")

    d1 = (Log(ForwardRate / Cap) + (Sigma ^ 2) * Time1 / 2) / Sigma * NRoot(Time1, 2)
    d2 = d1 - Sigma * NRoot(Time1, 2)

    Floorlet = -Loan * (Time2 - Time1) * (Exp(-InterestRate * Time2)) * ((ForwardRate * Excel.WorksheetFunction.NormDist(-d1, 0, 1, True)) - (Floor * Excel.WorksheetFunction.NormDist(-d2, 0, 1, True)))
    
    If Floorlet > 0 Then

    MsgBox ("The value of the Floorlet" & " " & Floorlet)

    ElseIf Floorlet <= 0 Then

    MsgBox ("The value of the Floorlet" & " " & 0)





End If





End Sub
