'Iterative solver for water depth in open channel
'Uses Manning Strickler flow formula

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
Private Function YnEval1(y As Double, Q As Double, Ks As Double, I As Double, b As Double, m As Double)
  YnEval1 = ((y * (b + m * y)) ^ (5 / 3) / _
    (b + 2 * y * (1 + m ^ 2) ^ 0.5) ^ (2 / 3)) - (Q / (Ks * I ^ 0.5))
End Function

'Evaluates f'(y)
Private Function YnPrimeEval1(y As Double, b As Double, m As Double)
  YnPrimeEval1 = (5 / 3) * (y * (b + m * y)) ^ (2 / 3) * (b + 2 * m * y) / _
    (b + 2 * y * (1 + m ^ 2) ^ 0.5) ^ (2 / 3) - (4 / 3) * _
    (y * (b + m * y)) ^ (5 / 3) * (1 + m ^ 2) ^ 0.5 / (b + 2 * _
    y * (1 + m ^ 2) ^ 0.5) ^ (5 / 3)
End Function

'Type II section
'Evaluates f(y)
Private Function YnEval2(y As Double, Q As Double, Ks As Double, I As Double, D As Double)
  YnEval2 = (1 / 64) * 8 ^ (1 / 3) * (D ^ 2 * (2 * WorksheetFunction.Acos(1 - 2 * y / D) - _
    Sin(2 * WorksheetFunction.Acos(1 - 2 * y / D)))) ^ (5 / 3) / (D * WorksheetFunction.Acos(1 - 2 * y / D)) ^ _
    (2 / 3) - Q / (Ks * I ^ 0.5)
End Function

'Evaluates f'(y)
Private Function YnPrimeEval2(y As Double, D As Double)
  YnPrimeEval2 = (5 / 192) * 8 ^ (1 / 3) * (D ^ 2 * (2 * WorksheetFunction.Acos(1 - 2 * y / D) - _
    Sin(2 * WorksheetFunction.Acos(1 - 2 * y / D)))) ^ (2 / 3) * D ^ 2 * (4 / (D * Sqr(1 - _
    (1 - 2 * y / D) ^ 2)) - 4 * Cos(2 * WorksheetFunction.Acos(1 - 2 * y / D)) / _
    (D * Sqr(1 - (1 - 2 * y / D) ^ 2))) / (D * WorksheetFunction.Acos(1 - 2 * y / D)) ^ (2 / 3) - _
    (1 / 48) * 8 ^ (1 / 3) * (D ^ 2 * (2 * WorksheetFunction.Acos(1 - 2 * y / D) - _
    Sin(2 * WorksheetFunction.Acos(1 - 2 * y / D)))) ^ (5 / 3) * D / ((D * WorksheetFunction.Acos(1 - 2 * _
    y / D)) ^ (5 / 3) * D * Sqr(1 - (1 - 2 * y / D) ^ 2))
End Function

'Normal Water Depth Problem Solving
'Uses Newton-Raphson method, 4th-order quadratic convergence

'Trapezoid sections
Function YNTRAPEZ(Q As Double, Ks As Double, I As Double, b As Double, m As Double)
  Dim y0 As Double
  Dim iter As Integer
  Dim oldy0 As Double
  
  y0 = INITIAL_SEED
  iter = 0
  
  Do
    oldy0 = y0
    y0 = y0 - YnEval1(y0, Q, Ks, I, b, m) / YnPrimeEval1(y0, b, m)
    iter = iter + 1
  Loop Until (Abs(y0 - oldy0) < ACCURACY) Or (iter > MAX_ITER)
  
  YNTRAPEZ = y0
End Function

'Rectangular sections
Function YNRECT(Q As Double, Ks As Double, I As Double, b As Double)
  YNRECT = YNTRAPEZ(Q, Ks, I, b, 0)
End Function

'Triangular sections
Function YNTRIANGLE(Q As Double, Ks As Double, I As Double, m As Double)
  YNTRIANGLE = YNTRAPEZ(Q, Ks, I, 0, m)
End Function

'Circular sections
Function YNCIRC(Q As Double, Ks As Double, I As Double, D As Double)
  Dim y0 As Double
  Dim iter As Integer
  Dim oldy0 As Double
  
  y0 = INITIAL_SEED
  iter = 0

  Do
    oldy0 = y0
    y0 = y0 - YnEval2(y0, Q, Ks, I, D) / YnPrimeEval2(y0, D)
    iter = iter + 1
  Loop Until (Abs(y0 - oldy0) < ACCURACY) Or (iter > MAX_ITER)
  
  YNCIRC = y0
End Function