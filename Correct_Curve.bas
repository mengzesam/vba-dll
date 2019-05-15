Attribute VB_Name = "Correct_Curve"
Public Function Exh_loss(annulus As Double) As Double

Exh_loss = 0.006919826 * annulus ^ 2 - 2.608241395 * annulus + 270.8153777
'-1.98714E-07    0.000527733 -0.413598849    109.1523082
'VV = annulus / 0.3048
'Exh_loss = 2.326 * (-0.000000198714 * VV ^ 3 + 0.000527733 * VV ^ 2 - 0.413598849 * VV + 109.1523082)
Exh_loss = 25.56 * (annulus / 177) ^ 1 - 3
End Function

