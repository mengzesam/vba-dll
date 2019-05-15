Attribute VB_Name = "Flow_Discharge_Coefficient"
Function Cfun(mytype As String, Red As Double, beta As Double, Optional Cx As Double = 1#, Optional tapping As Integer = 1, Optional d As Double = 500#) As Double
If (mytype = "³¤¾±Åç×ì") Then
  Cfun = C_LongradiusNozzle(Red)
ElseIf (mytype = "ISA32Åç×ì") Then
  Cfun = C_ISA32Nozzle(Red, beta)
ElseIf (mytype = "ASMEÅç×ì") Then
  Cfun = C_ASMENozzle(Red, Cx)
ElseIf (mytype = "¿×°åStolz") Then
  Cfun = C_Stolz(Red, beta, tapping, d)
End If
End Function

Function C_LongradiusNozzle(Red As Double) As Double

C_LongradiusNozzle = 0.9965 - 0.00653 * (1000000# / Red) ^ 0.5

End Function


Function C_ISA32Nozzle(Red As Double, beta As Double) As Double

C_ISA32Nozzle = 0.99 - 0.2262 * beta ^ 4.1 - (0.00175 * beta ^ 2 - 0.0033 * beta ^ 4.15) * (1000000# / Red / beta) ^ 1.15

End Function

Function C_ASMENozzle(Red As Double, Cx As Double) As Double
 C_ASMENozzle = Cx - 0.185 * Red ^ (-0.2) * (1 - 361239 / Red) ^ 0.8
End Function

Function C_Stolz(Red As Double, beta As Double, tapping As Integer, d As Double) As Double  'orifice by Stolz eq.
Dim L1 As Double, L2 As Double, K As Double
'tapping=1:corner tapping
'tapping=2:D and D/2 tapping
'tapping=3:flange tapping
 C_Stolz = 0.5959 + 0.0312 * beta ^ 2.1 - 0.184 * beta ^ 8 + 0.0029 * beta ^ 2.5 * (1000000# / Red / beta) ^ 0.75
 L1 = 0
 L2 = 0
 K = 0.09
 If (tapping = 2) Then
   L1 = 1
   L2 = 0.47
   K = 0.039
 ElseIf (tapping = 3) Then
   L1 = 25.4 / d
   L2 = L1
   If (L1 >= 0.4333) Then
   K = 0.039
   End If
 End If
 C_Stolz = C_Stolz + K * L1 * beta ^ 4 / (1 - beta ^ 4) - 0.0337 * L2 * beta ^ 3
End Function

