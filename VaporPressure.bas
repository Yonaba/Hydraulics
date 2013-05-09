Attribute VB_Name = "VaporPressure"

'Evaluates the vapor pressure for water for pure water
'Constants taken from A.L. Mar (2003) - Cours d'Hydraulique en Charge

'Interpolation
Private Function InterpolatePPEP(PPEP, Temp As Double) As Double
  Dim TempH, TempL As Integer
  TempH = WorksheetFunction.Ceiling(Temp, 5)
  TempL = TempH - 5
  InterpolatePPEP = ((Temp - TempL) * (PPEP(1 + (TempH / 5)) - PPEP(1 + (TempL / 5))) / 10) + PPEP(1 + (TempL / 5))
End Function

'Returns the vapor pressure for pure water
'Arg temp represents the temperature in celsius degrees
'Temp. range supported (as of now) goes from 0 to 100°
Function VaporPressure(Temp As Double) As Double
  Dim PPEP(1 To 21) As Double
  PPEP(1) = 0.06
  PPEP(2) = 0.09
  PPEP(3) = 0.12
  PPEP(4) = 0.17
  PPEP(5) = 0.25
  PPEP(6) = 0.33
  PPEP(7) = 0.44
  PPEP(8) = 0.58
  PPEP(9) = 0.76
  PPEP(10) = 1.01
  PPEP(11) = 1.26
  PPEP(12) = 1.61
  PPEP(13) = 2.03
  PPEP(14) = 2.56
  PPEP(15) = 3.2
  PPEP(16) = 3.96
  PPEP(17) = 4.86
  PPEP(18) = 5.93
  PPEP(19) = 7.18
  PPEP(20) = 8.62
  PPEP(21) = 10.33
  If Temp >= 0 And Temp <= 100 Then
    If (Temp Mod 5) = 0 Then
      VaporPressure = PPEP((Temp / 5) + 1)
      Exit Function
    Else
      VaporPressure = InterpolatePPEP(PPEP, Temp)
    End If
  Else
    VaporPressure = vbNullString
  End If
End Function