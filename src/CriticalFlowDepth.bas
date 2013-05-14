Attribute VB_Name = "CriticalDepth"

'Iterative solver for critical water depth in open channel

'Force type declaration
Option Explicit

'Iteration control constants
Private Const ACCURACY = 0.000000001
Private Const INITIAL_SEED = 0.1
Private Const MAX_ITER = 100

'Parametrized sections supported so far:
' Type I:
'   * Trapezoid : b, m, y
'   * Rectangular (b, m = 0, y)
'   * Triangular (b = 0, m, y)
' Type II:
' * Circular (d, y)

'Type I section
'Evaluates f(y)
Private Function YcEval1(y As Double, Q As Double, b As Double, m1 As Double, m2 As Variant, g As Variant)
  YcEval1 = (1 / 8) * (2 * b + y * ((1 + m1 ^ 2) ^ 0.5 + (1 + m2 ^ 2) ^ 0.5)) ^ _
    3 * y ^ 3 / (b + y * (m1 + m2)) - Q ^ 2 / g
End Function

'Evaluates f'(y)
Private Function YcPrimeEval1(y As Double, b As Double, m1 As Double, m2 As Variant)
  YcPrimeEval1 = (3 / 8) * (2 * b + y * ((1 + m1 ^ 2) ^ 0.5 + (1 + m2 ^ 2) ^ 0.5)) ^ _
    2 * y ^ 3 * ((1 + m1 ^ 2) ^ 0.5 + (1 + m2 ^ 2) ^ 0.5) / (b + y * (m1 + m2)) + _
    (3 / 8) * (2 * b + y * ((1 + m1 ^ 2) ^ 0.5 + (1 + m2 ^ 2) ^ 0.5)) ^ 3 * y ^ 2 / _
    (b + y * (m1 + m2)) - (1 / 8) * (2 * b + y * ((1 + m1 ^ 2) ^ 0.5 + (1 + m2 ^ 2) ^ 0.5)) ^ _
    3 * y ^ 3 * (m1 + m2) / (b + y * (m1 + m2)) ^ 2
End Function

'Type II section
'Evaluates f(y)
Private Function YcEval2(t As Double, Q As Double, D As Double, g As Double)
  YcEval2 = ((1 / 8) * D ^ 2 * (t - Sin(t))) ^ 3 / (D * Sin((1 / 2) * t)) - Q ^ 2 / g
End Function

'Evaluates f'(y)
Private Function YcPrimeEval2(t As Double, D As Double)
  YcPrimeEval2 = (3 / 512) * D ^ 5 * (t - Sin(t)) ^ 2 * (1 - Cos(t)) / _
    Sin((1 / 2) * t) - (1 / 1024) * D ^ 5 * (t - Sin(t)) ^ 3 * Cos((1 / 2) * t) / _
    Sin((1 / 2) * t) ^ 2
End Function

'Critical water depth solver
'Uses Newton-Raphson method, 4th-order quadratic convergence

'Trapezoid sections
Function YCTRAPEZ(Q As Double, b As Double, m1 As Double, Optional m2 As Variant, Optional g As Variant)
  Dim yc As Double
  Dim oldyc As Double
  Dim iter As Integer
  
  If IsMissing(m2) Then m2 = m1
  If IsMissing(g) Then g = 9.81

  yc = INITIAL_SEED
  Do
    oldyc = yc
    yc = yc - YcEval1(yc, Q, b, m1, m2, g) / YcPrimeEval1(yc, b, m1, m2)
    iter = (iter Or 0) + 1
  Loop Until (Abs(yc - oldyc) < ACCURACY) Or (iter > MAX_ITER)
  
  YCTRAPEZ = yc
End Function

'Rectangular sections
Function YCRECT(Q As Double, b As Double, Optional g As Double = 9.81)
  YCRECT = (Q / (b * g ^ 0.5)) ^ (2 / 3)
End Function

'Triangular sections
Function YCTRIANGLE(Q As Double, m1 As Double, m2 As Variant, Optional g As Variant)
  If IsMissing(m2) Then m2 = m1
  If IsMissing(g) Then g = 9.81
  If m1 = m2 Then
    YCTRIANGLE = ((2 * Q ^ 2) / ((m1 ^ 2) * g)) ^ (1 / 5)
  Else
    YCTRIANGLE = ((2 ^ 3 * Q ^ 2) / (g * (m1 + m2) ^ 2)) ^ (1 / 5)
  End If
End Function

'Circular sections
'Solving the problem here appears quite complex as it results in an erratic behavior
'from the solver. So we need to start iterating from a seed close to the actual answer.
'The approximation of the critical water depth used here was taken from the paper
'HSL, J. Vasquez, ENGEES - 2010 (p.80)
'Available online at: http://engees.unistra.fr/site/fileadmin/user_upload/pdf/shu/cours_HSL_FI_2010.pdf
Function YCCIRC(Q As Double, D As Double, Optional g As Double = 9.81)
  Dim oldTheta As Double
  Dim theta As Double
  Dim yc As Double
  Dim iter As Integer
  
  'Approximation of the critical depth, to work out a good seed for theta
  yc = (Q / (Q * D) ^ 0.5) ^ 0.5
  theta = 2 * WorksheetFunction.Acos(1 - (2 * yc / D))

  Do
    oldTheta = theta
    theta = theta - YcEval2(theta, Q, D, g) / YcPrimeEval2(theta, D)
    iter = (iter Or 0) + 1
  Loop Until (Abs(theta - oldTheta) < ACCURACY) Or (iter > MAX_ITER)
  
  YCCIRC = (D / 2) * (1 - Cos(theta / 2))
End Function