'Iterative solver for open channel flow water depth evaluation
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

'Normal Water Depth Problem Solving
'Uses Newton-Raphson method, 4th-order quadratic convergence

'Type I sections
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

Function YNRECT(Q As Double, Ks As Double, I As Double, b As Double)
  YNRECT = YNTRAPEZ(Q, Ks, I, b, 0)
End Function

Function YNTRIANGLE(Q As Double, Ks As Double, I As Double, m As Double)
  YNTRIANGLE = YNTRAPEZ(Q, Ks, I, 0, m)
End Function