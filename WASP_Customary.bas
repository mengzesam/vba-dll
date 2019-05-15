Attribute VB_Name = "WASP_Customary"
'Option Explicit

'Make sure Excel can find steam67.dll by putting it in the windows directory.
'Else specify full path here for dll.
'The dimensional units are Temperature - F, Pressure - psia, Steam Quality - %,
'Specific Volume - cuft/lbm, Specific Enthalpy - btu/lbm, Specific Entropy - btu/lbm/R,
'Saturation Temperature - F, Saturation Pressure - psia, Degrees of Superheat - F,
'Degrees of Subcooling - F, Viscosity - lbf*s/ft^2 and Critical Velocity - ft/s.

Declare Function steam67 Lib "steam67.dll" _
                                (temperature As Double, pressure As Double, _
                                 quality As Double, weight As Double, _
                                 enthalpy As Double, entropy As Double, _
                                 saturation_temperature As Double, saturation_pressure As Double, _
                                 degrees_superheat As Double, degrees_subcooling As Double, _
                                 viscosity As Double, critical_velocity As Double, _
                                 ByVal action As Long) As Long

Function EN_H_PT(P, T) As Double
Dim X As Double, W As Double, H As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_H_PT = H
End Function

Function EN_S_PT(P, T) As Double
Dim X As Double, W As Double, H As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_S_PT = S
End Function

Function EN_Ts_Ps(Ps) As Double
Dim T As Double, X As Double, W As Double, H As Double, S As Double, Ts As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, Ps, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_Ts_Ps = Ts
End Function

Function EN_Ps_Ts(Ts) As Double
Dim P As Double, X As Double, W As Double, H As Double, S As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(Ts, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_Ps_Ts = Ps
End Function

Function EN_H_PS(P, S) As Double
Dim T As Double, X As Double, W As Double, H As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_H_PS = H
End Function

Function EN_S_PH(P, H) As Double
Dim T As Double, X As Double, W As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_S_PH = S
End Function

Function EN_X_PH(P, H) As Double
Dim T As Double, X As Double, W As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_X_PH = X
End Function

Function EN_v_PH(P, H) As Double
Dim T As Double, X As Double, W As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_v_PH = W
End Function

Function EN_T_PH(P, H) As Double
Dim T As Double, X As Double, W As Double, S As Double, Ts As Double, Ps As Double
Dim supHt As Double, subCl As Double, Vi As Double, VELcr As Double, Act As Long, lrt As Long
Act = 0
lrt = steam67(T, P, X, W, H, S, Ts, Ps, supHt, subCl, Vi, VELcr, Act)
EN_T_PH = T
End Function

