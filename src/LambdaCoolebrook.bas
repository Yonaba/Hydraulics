Attribute VB_Name = "LambdaCoolebrook"

'Iterative solver for lambda head-loss ratio
'Required for Darcy-Weisbach head-loss calculation

'Force type declaration
Option Explicit

'Internal constants
Private Const ACCURACY = 0.000000001
Private Const WATER_VISCOSITY = 1 * (10 ^ (-6))
Private Const MAX_ITER = 100
Private Const DEF_BFRONTIER = 2300
Private Const PI = 3.14159265358979

'As VB has no in-built function to to compute log10, we provide support for it
Static Function Log10(X As Double) As Double
  Log10 = Log(X) / Log(10#)
End Function

'Lambda evaluation for Poiseuille flow
Private Function fPoiseuille(Re As Double) As Double
  fPoiseuille = 64 / Re
End Function

'Reynolds Number

Public Function Reynolds(Q As Double, D As Double, Optional nu As Double = WATER_VISCOSITY) As Double
        Reynolds = (4 * Q) / (PI * D * nu)
End Function
 
'Lambda evaluation using Swamee & Jain approximation
Public Function Lambda_SwameeJain(k As Double, D As Double, Q As Double, Optional nu As Double = WATER_VISCOSITY) As Double
        Dim Re As Double
        Re = Reynolds(Q, D, nu)
        Lambda_SwameeJain = 0.25 / ((Log10((k / (3.71 * D)) + (5.74 / (Re ^ 0.9))))) ^ 2
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
'Reverts to Laminar Poiseuille's flow for very low Reynolds values
'Uses Swamee-Jain approximation for turbulent flow as a seed for iterations
'Convergence is 4th-order quadratic
Public Function Lambda(k As Double, D As Double, Q As Double, Optional nu As Double = WATER_VISCOSITY, Optional bFrontier As Double = DEF_BFRONTIER, Optional maxIter = MAX_ITER) As Double
        Dim x0 As Double, oldx0 As Double
        Dim i As Integer
        Dim Re As Double
        Dim ll As Double
        
        Re = Reynolds(Q, D, nu)
        
        If Re < bFrontier Then
                ll = fPoiseuille(Re)
        Else
                x0 = Lambda_SwameeJain(k, D, Q, nu)
                i = 0
                Do
                        oldx0 = x0
                        x0 = x0 - fEval(x0, k, D, Re) / fPrimeEval(x0, k, D, Re)
                        i = i + 1
                Loop Until ((Abs(x0 - oldx0) < ACCURACY) Or (i > maxIter))
                ll = x0
        End If
        Lambda = ll
End Function
