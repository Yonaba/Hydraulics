'Iterative solver for open channel flow water depth evaluation
'Uses Manning Strickler flow formula

'Force type declaration
Option Explicit

Private Const ACCURACY = 0.000000001
Private Const INITIAL_SEED = 1
Private Const MAX_ITER = 100

'Manning Strickler derivative
Private Function YnPrimeEval(y As Double, b As Double, m As Double)
	YnPrimeEval = (5 / 3) * (y * (b + m * y)) ^ (2 / 3) * (b + 2 * m * y) / _
								(b + 2 * y * (1 + m ^ 2) ^ 0.5) ^ (2 / 3) - (4 / 3) * _
								(y * (b + m * y)) ^ (5 / 3) * (1 + m ^ 2) ^ 0.5 / (b + 2 * _
								y * (1 + m ^ 2) ^ 0.5) ^ (5 / 3)
End Function

'Manning Strickler function evaluation
Private Function YnEval(y As Double, Q As Double, Ks As Double, I As Double, b As Double, m As Double)
	YnEval = ((y * (b + m * y)) ^ (5 / 3) / _
					 (b + 2 * y * (1 + m ^ 2) ^ 0.5) ^ (2 / 3)) - (Q / (Ks * I ^ 0.5))
End Function

'Water depth solver
'Uses Newton-Raphson method, 4th-order quadratic convergence
Function YN(Q As Double, Ks As Double, I As Double, b As Double, m As Double)
	Dim y0 As Double
	Dim iter As Integer
	Dim oldy0 As Double
	
	y0 = INITIAL_SEED
	iter = 0
	
	Do
			oldy0 = y0
			y0 = y0 - YnEval(y0, Q, Ks, I, b, m) / YnPrimeEval(y0, b, m)
			iter = iter + 1
	Loop Until (Abs(y0 - oldy0) < ACCURACY) Or (iter > MAX_ITER)
	
	Yn = y0
End Function