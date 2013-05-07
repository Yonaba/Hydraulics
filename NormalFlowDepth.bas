Attribute VB_Name = "NormalDepth"

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
Private Function YnEval2(t As Double, Q As Double, Ks As Double, I As Double, D As Double)
  YnEval2 = (1 / 64) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ _
    (5 / 3) * 2 ^ (2 / 3) / (D * t) ^ (2 / 3) - Q / (Ks * I ^ 0.5)
  
End Function

'Evaluates f'(y)
Private Function YnPrimeEval2(t As Double, D As Double)
  YnPrimeEval2 = (5 / 192) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ _
    (2 / 3) * 2 ^ (2 / 3) * D ^ 2 * (1 - Cos(t)) / (D * t) ^ (2 / 3) - _
    (1 / 96) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ (5 / 3) * 2 ^ _
    (2 / 3) * D / (D * t) ^ (5 / 3)
End Function

'Normal Water Depth Problem Solving
'Uses Newton-Raphson method, 4th-order quadratic convergence

'Trapezoid sections
Function YNTRAPEZ(Q As Double, Ks As Double, I As Double, b As Double, m As Double)
  Dim yn As Double
  Dim oldyn As Double
  Dim iter As Integer
  
  yn = INITIAL_SEED
  
  Do
    oldyn = yn
    yn = yn - YnEval1(yn, Q, Ks, I, b, m) / YnPrimeEval1(yn, b, m)
    iter = (iter Or 0) + 1
  Loop Until (Abs(yn - oldyn) < ACCURACY) Or (iter > MAX_ITER)
  
  YNTRAPEZ = yn
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
'In order to avoid an erratic behavior from the solver, we also need  here (again) a good seed
'We will use then an approx. of the critical water as a seed to workout the water normal depth
'See the implementation CriticalFlowDepth.bas for more details
Function YNCIRC(Q As Double, Ks As Double, I As Double, D As Double)
  Dim yn As Double
  Dim iter As Integer
  Dim theta As Double
  Dim oldTheta As Double
  
  yn = (Q / (Q * D) ^ 0.5) ^ 0.5
  theta = 2 * WorksheetFunction.Acos(1 - (2 * yn / D))
  
  Do
    oldTheta = theta
    theta = theta - YnEval2(theta, Q, Ks, I, D) / YnPrimeEval2(theta, D)
    iter = (iter Or 0) + 1
  Loop Until (Abs(theta - oldTheta) < ACCURACY) Or (iter > MAX_ITER)
  
  YNCIRC = (D / 2) * (1 - Cos(theta / 2))
End Function