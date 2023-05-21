Attribute VB_Name = "NormalDepth"

'Iterative solver for water depth in open channel
'Uses Manning Strickler flow formula

'Force type declaration
Option Explicit

'Iteration control constants
Private Const ACCURACY = 0.000000001
Private Const INITIAL_SEED = 0.1
Private Const MAX_ITER = 100
Private Const PI = 3.14159265358979

Private Function NDPDebitance(Q As Double, Ks As Double, I As Double) As Double
  NDPDebitance = (Q / (Ks * (I ^ 0.5)))
End Function

Private Function TrapezeSection(y As Double, b As Double, m As Double, n As Double) As Double
  TrapezeSection = y * (b + (1 / 2 * (m + n)) * y)
End Function

Private Function TrapezeWetPerimeter(y As Double, b As Double, m As Double, n As Double) As Double
  TrapezeWetPerimeter = b + y * ((1 + m ^ 2) ^ 0.5 + (1 + n ^ 2) ^ 0.5)
End Function

Private Function NDPTrapezeSectionDrv(y As Double, b As Double, m As Double, n As Double) As Double
  NDPTrapezeSectionDrv = (5 / 3) * (y * (b + (1 / 2 * (m + n)) * y)) ^ (2 / 3) * _
    (b + (1 / 2 * (m + n)) * y + ((1 / 2) * m + (1 / 2) * n) * y)
End Function

Private Function NPDTrapWetPerimDrv(y As Double, b As Double, m As Double, n As Double) As Double
  NPDTrapWetPerimDrv = (((1 + m ^ 2) ^ 0.5 + (1 + n ^ 2) ^ 0.5) / ((b + y * ((1 + m ^ 2) ^ 0.5 + _
    (1 + n ^ 2) ^ 0.5)) ^ (1 / 3))) * 2 / 3
End Function

Private Function NDPFnTrapezeEval(y As Double, Q As Double, Ks As Double, I As Double, b As Double, m As Double, n As Double) As Double
  NDPFnTrapezeEval = ((TrapezeSection(y, b, m, n) ^ (5 / 3)) / (TrapezeWetPerimeter(y, b, m, n) ^ (2 / 3))) - NDPDebitance(Q, Ks, I)
End Function

Private Function NDPFnPrimeTrapezeEval(y As Double, b As Double, m As Double, n As Double) As Double
  Dim Sect As Double
  Dim Perim As Double
  Dim SectDrv As Double
  Dim PerimDrv As Double
  
  Sect = (TrapezeSection(y, b, m, n)) ^ (5 / 3)
  SectDrv = NDPTrapezeSectionDrv(y, b, m, n)
  Perim = (TrapezeWetPerimeter(y, b, m, n)) ^ (2 / 3)
  PerimDrv = NPDTrapWetPerimDrv(y, b, m, n)
  NDPFnPrimeTrapezeEval = (SectDrv * Perim - PerimDrv * Sect) / (Perim ^ 2)
End Function

Private Function NDPEvalCircTheta(t As Double, Q As Double, Ks As Double, I As Double, D As Double)
  NDPEvalCircTheta = (1 / 64) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ _
    (5 / 3) * 2 ^ (2 / 3) / (D * t) ^ (2 / 3) - Q / (Ks * I ^ 0.5)
End Function

Private Function NDPPrimeEvalCircTheta(t As Double, D As Double)
  NDPPrimeEvalCircTheta = (5 / 192) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ _
    (2 / 3) * 2 ^ (2 / 3) * D ^ 2 * (1 - Cos(t)) / (D * t) ^ (2 / 3) - _
    (1 / 96) * 8 ^ (1 / 3) * (D ^ 2 * (t - Sin(t))) ^ (5 / 3) * 2 ^ _
    (2 / 3) * D / (D * t) ^ (5 / 3)
End Function

Function YNTRAPEZE(Q As Double, Ks As Double, I As Double, b As Double, m1 As Double, m2 As Double) As Double
  Dim Yn As Double
  Dim OldYn As Double
  Dim countIter As Integer
  Yn = INITIAL_SEED
  
  Do
    OldYn = Yn
    Yn = Yn - NDPFnTrapezeEval(Yn, Q, Ks, I, b, m1, m2) / NDPFnPrimeTrapezeEval(Yn, b, m1, m2)
    countIter = (countIter Or 0) + 1
  Loop Until (Abs(Yn - OldYn) < ACCURACY) Or (countIter > MAX_ITER)
  
  YNTRAPEZE = Yn
End Function

Function YNTRAPEZEISO(Q As Double, Ks As Double, I As Double, b As Double, m1 As Double) As Double
  YNTRAPEZEISO = YNTRAPEZE(Q, Ks, I, b, m1, m1)
End Function

Function YNTRIANGLE(Q As Double, Ks As Double, I As Double, m1 As Double, m2 As Double) As Double
  YNTRIANGLE = YNTRAPEZE(Q, Ks, I, 0, m1, m2)
End Function

Function YNTRIANGLEISO(Q As Double, Ks As Double, I As Double, m1 As Double) As Double
  YNTRIANGLEISO = YNTRAPEZE(Q, Ks, I, 0, m1, m1)
End Function

Function YNRECTANGLE(Q As Double, Ks As Double, I As Double, b As Double) As Double
  YNRECTANGLE = YNTRAPEZE(Q, Ks, I, b, 0, 0)
End Function

Function YNCIRCULAR(Q As Double, Ks As Double, I As Double, D As Double)
  Dim qn As Double, Yn As Double
  Dim theTa As Double
  Dim oldTheta As Double
  Dim countIter As Integer
  
  qn = Q / (Ks * (I ^ 0.5) * (D ^ (8 / 3)))
  Yn = ((11 * D) / (5 * PI)) * WorksheetFunction.Asin(1.614 * (qn ^ 0.485))
  theTa = 2 * WorksheetFunction.Acos(1 - (2 * Yn / D))
  
  Do
    oldTheta = theTa
    theTa = theTa - NDPEvalCircTheta(theTa, Q, Ks, I, D) / NDPPrimeEvalCircTheta(theTa, D)
    countIter = (countIter Or 0) + 1
  Loop Until (Abs(theTa - oldTheta) < ACCURACY) Or (countIter > MAX_ITER)
  
  YNCIRCULAR = (D / 2) * (1 - Cos(theTa / 2))
End Function
