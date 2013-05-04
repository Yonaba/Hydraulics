'Iterative solver for open channel critical water flow depth evaluation

'Force type declaration
Option Explicit

Private Const ACCURACY = 0.000000001
Private Const INITIAL_SEED = 0.1
Private Const MAX_ITER = 100

'Manning Strickler derivative
Private Function PrimeEval(y As Double, b As Double, m As Double)
    PrimeEval = 3 * y ^ 2 * (b + m * y) ^ 3 / (b + 2 * m * y) + 3 * y ^ 3 * (b + m * y) ^ 2 * m / _
                (b + 2 * m * y) - 2 * y ^ 3 * (b + m * y) ^ 3 * m / (b + 2 * m * y) ^ 2
End Function

'Manning Strickler function evaluation
Private Function YcEqEval(y As Double, Q As Double, b As Double, m As Double, g As Double)
    YcEqEval = y ^ 3 * (b + m * y) ^ 3 / (b + 2 * m * y) - (Q ^ 2 / g)
End Function

'Water depth solver
'Uses Newton-Raphson method, 4th-order quadratic convergence
Function YC(Q As Double, b As Double, m As Double, Optional g As Double = 9.81)
    Dim y0 As Double
    Dim iter As Integer
    Dim oldy0 As Double
    
    y0 = INITIAL_SEED
    iter = 0
    
    Do
            oldy0 = y0
            y0 = y0 - YcEqEval(y0, Q, b, m, g) / PrimeEval(y0, b, m)
            iter = iter + 1
    Loop Until (Abs(y0 - oldy0) < ACCURACY) Or (iter > MAX_ITER)
    
    YC = y0
End Function


