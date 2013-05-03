'Iterative solver for lambda head-loss ratio
'Required for Darcy-Weisbach head-loss calculation

'Force type declaration
Option Explicit

'Internal constants
Private Const ACCURACY = 0.000000001
Private Const INITIAL_SEED = 0.01
Private Const MAX_ITER = 100
Private Const DEF_BFRONTIER = 2300

'As VB has no in-built function to to compute log10, we provide support for it
Static Function Log10(X As Double) As Double
  Log10 = Log(X) / Log(10#)
End Function

'Lambda evaluation for Poiseuille flow
Private Function fPoiseuille(Re As Double) As Double
  fPoiseuille = 64 / Re
End Function

'Coolebrook-White lambda derivative
Private Function fPrimeEval(xn As Double, k As Double, D As Double, Re As Double) As Double
  fPrimeEval = (-1# / (2# * xn ^ (1.5)) - (2.51 / (Re * xn ^ 1.5 * Log(10#) * ((2.51 / (Re * xn ^ 0.5)) + (0.269541779 * k / D)))))
End Function

'Coolebrook-White lambda evaluation
Private Function fEval(xn As Double, k As Double, D As Double, Re As Double) As Double
  fEval = ((1# / (xn ^ 0.5)) + (2# * Log10((k / (3.71 * D)) + (2.51 / (Re * (xn ^ 0.5))))))
End Function

'Coolebrook-White lambda solving using Newton-Raphson method
'Convergence is 4th-order quadratic
Public Function LAMBDA(k As Double, D As Double, Re As Double, Optional bFrontier As Double = DEF_BFRONTIER, Optional maxIter = MAX_ITER) As Double
	Dim x0, oldx0 As Double
  Dim i As Integer
	
	If Re < bFrontier Then
		LAMBDA = fPoiseuille(Re)
	End If
	x0 = INITIAL_SEED
	
	i = 0
	Do
		oldx0 = x0
		x0 = x0 - fEval(x0, k, D, Re) / fPrimeEval(x0, k, D, Re)
		i = i + 1
	Loop Until ((Abs(x0 - oldx0) < ACCURACY) Or (i > maxIter))
	
	LAMBDA = x0
End Function