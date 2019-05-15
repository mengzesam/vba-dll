Attribute VB_Name = "Expand_Line"
Option Explicit
Option Base 1

'funtype=0:straight line to fitting, connecting line from the first node to the last
'funtype=1 :Bezier curve to fitting
'funtype=2: Polynomial curve to fitting
Function seekCrossInHS(P As Double, Known_Y As Range, Known_X As Range, Optional funtype As Integer = 1, _
Optional stnode As Integer = 1, Optional endnode As Integer = 2, Optional polyOrder As Integer = 2) As Variant
 If funtype = 2 Then
   seekCrossInHS = seekCrossInHS_Poly(P, Known_Y, Known_X, 2)
 ElseIf funtype = 1 Then
   seekCrossInHS = seekCrossInHS_Bezier(P, Known_Y, Known_X)
 ElseIf funtype = 0 Then
   seekCrossInHS = seekCrossInHS_Line(P, Known_Y, Known_X, stnode, endnode)
 End If
End Function
 
 Function seekCrossInHS_Line(P As Double, Known_Y As Range, Known_X As Range, Optional stnode As Integer = 1, Optional endnode As Integer = 2) As Variant
  Dim delta As Double, K As Integer
  Dim S1 As Double, S2 As Double, S As Double
  Dim H1 As Double, H2 As Double, H As Double
  Dim YY(2) As Double, XX(2) As Double
  Dim polys As Variant
  
  On Error GoTo errLab
  If Known_X.Count <> Known_Y.Count Then
    seekCrossInHS_Line = "X size<>Y size"
    Exit Function
  End If
  If Known_X.Count < 2 Then
    seekCrossInHS_Line = "Size<2"
    Exit Function
  End If
  If stnode >= endnode Or stnode < 1 Or endnode > Known_X.Count Then
    seekCrossInHS_Line = "Error start/end node"
  End If
  
  delta = 0.0000001
   S1 = Known_X(stnode).Value
   S2 = Known_X(endnode).Value
   H1 = Known_Y(stnode).Value
   H2 = Known_Y(endnode).Value
   XX(1) = S1
   XX(2) = S2
   YY(1) = H1
   YY(2) = H2
   polys = Application.WorksheetFunction.LinEst(YY, XX)
   H1 = polys(1) * S1 + polys(2) - H_PS(P, S1)
   H2 = polys(1) * S2 + polys(2) - H_PS(P, S2)
   S = S1 - H1 * (S1 - S2) / (H1 - H2)
   H = polys(1) * S + polys(2) - H_PS(P, S)
   K = 0
Do While (Abs(H) > delta And K < 100)
   K = K + 1
   S2 = S1
   H2 = H1
   S1 = S
   H1 = H
   S = S1 - H1 * (S1 - S2) / (H1 - H2)
   H = polys(1) * S + polys(2) - H_PS(P, S)
Loop
If Not K < 100 Then
 seekCrossInHS_Line = "No Solving or Multi-Solving"
End If
seekCrossInHS_Line = H_PS(P, S)
Exit Function
  
errLab:
    seekCrossInHS_Line = "No Solving"
 End Function
 
 Function seekCrossInHS_Bezier(P As Double, Known_Y As Range, Known_X As Range) As Variant
  Dim delta As Double, K As Integer
  Dim S1 As Double, S2 As Double, S As Double
  Dim H1 As Double, H2 As Double, H As Double
  Dim rs As Variant
  
  On Error GoTo errLab
  
  If Known_X.Count <> Known_Y.Count Then
    seekCrossInHS_Bezier = "X size<>Y size"
    Exit Function
  End If
  If Known_X.Count < 2 Then
    seekCrossInHS_Bezier = "Size<2"
    Exit Function
  End If
delta = 0.0000001
H1 = Application.WorksheetFunction.Min(Known_Y) + 0.001
H2 = Application.WorksheetFunction.Max(Known_Y) - 0.001

S1 = Application.WorksheetFunction.Index(BezierFit(Known_X, Known_Y, H1, 1, "y"), 1) - S_PH(P, H1)
S2 = Application.WorksheetFunction.Index(BezierFit(Known_X, Known_Y, H2, 1, "y"), 1) - S_PH(P, H2)
If (S1 < delta) Then
  seekCrossInHS_Bezier = seekCrossInHS_Line(P, Known_Y, Known_X, Known_X.Count - 1, Known_X.Count)
   Exit Function
ElseIf (S2 > delta) Then
 seekCrossInHS_Bezier = seekCrossInHS_Line(P, Known_Y, Known_X, 1, 2)
  Exit Function
End If
H = H1 - S1 * (H1 - H2) / (S1 - S2)
If (H < Application.WorksheetFunction.Min(Known_Y) + 0.001) Then
     H = Application.WorksheetFunction.Min(Known_Y) + 0.001
   ElseIf (H > Application.WorksheetFunction.Max(Known_Y) - 0.001) Then
     H = Application.WorksheetFunction.Max(Known_Y) - 0.001
End If
S = Application.WorksheetFunction.Index(BezierFit(Known_X, Known_Y, H, 1, "y"), 1) - S_PH(P, H)
K = 0
Do While (Abs(S) > delta And K < 100)
     
   K = K + 1
   S2 = S1
   H2 = H1
   S1 = S
   H1 = H
   If (Abs(H1 - H2) < delta) Then
    H = H1
    Exit Do
   End If
   H = H1 - S1 * (H1 - H2) / (S1 - S2)
   If (H < Application.WorksheetFunction.Min(Known_Y) + 0.001) Then
     H = Application.WorksheetFunction.Min(Known_Y) + 0.001
   ElseIf (H > Application.WorksheetFunction.Max(Known_Y) - 0.001) Then
     H = Application.WorksheetFunction.Max(Known_Y) - 0.001
   End If
     
   S = Application.WorksheetFunction.Index(BezierFit(Known_X, Known_Y, H, 1, "y"), 1) - S_PH(P, H)
Loop
If Not K < 100 Then
 seekCrossInHS_Bezier = "No Solving or Multi-Solving"
 Exit Function
End If
seekCrossInHS_Bezier = H
Exit Function
errLab:
    seekCrossInHS_Bezier = "No Solving"
 End Function
 
 
Function seekCrossInHS_Poly(P As Double, Known_Y As Range, Known_X As Range, Optional polyOrder As Integer = 2) As Variant
  Dim delta As Double, i As Integer, j As Integer, K As Integer
  Dim S1 As Double, S2 As Double, S As Double
  Dim H1 As Double, H2 As Double, H As Double
  Dim polys As Variant
  Dim YS() As Double, XS() As Double
  Dim SS() As Double

  On Error GoTo errLab
  If Known_X.Count <> Known_Y.Count Then
    seekCrossInHS_Poly = "X size<>Y size"
    Exit Function
  End If
  If Known_X.Count < 2 Then
    seekCrossInHS_Poly = "Size<2"
    Exit Function
  End If
  ReDim YS(1 To Known_Y.Count)
  ReDim XS(1 To polyOrder + 1, 1 To Known_X.Count)
  For i = 1 To Known_X.Count
      YS(i) = Known_Y(i).Value
      For j = 1 To polyOrder + 1
       XS(j, i) = Known_X(i).Value ^ (j - 1)
      Next
  Next
    polys = Application.WorksheetFunction.LinEst(YS, XS, False, False)
        
delta = 0.00001
S1 = Known_X(1).Value
S2 = Known_X(Known_X.Count).Value
ReDim SS(1 To polyOrder + 2)
For i = 1 To polyOrder + 2
  SS(i) = S1 ^ (polyOrder + 1 - i)
Next
H1 = Application.WorksheetFunction.SumProduct(polys, SS) - H_PS(P, S1)
For i = 1 To polyOrder + 2
  SS(i) = S2 ^ (polyOrder + 1 - i)
Next
H2 = Application.WorksheetFunction.SumProduct(polys, SS) - H_PS(P, S2)
S = S1 - H1 * (S1 - S2) / (H1 - H2)
For i = 1 To polyOrder + 2
  SS(i) = S ^ (polyOrder + 1 - i)
Next
H = Application.WorksheetFunction.SumProduct(polys, SS) - H_PS(P, S)
K = 0
Do While (Abs(H) > delta And K < 100)
   K = K + 1
   S2 = S1
   H2 = H1
   S1 = S
   H1 = H
   S = S1 - H1 * (S1 - S2) / (H1 - H2)
   For i = 1 To polyOrder + 2
     SS(i) = S ^ (polyOrder + 1 - i)
   Next
   H = Application.WorksheetFunction.SumProduct(polys, SS) - H_PS(P, S)
Loop
If Not K < 100 Then
 seekCrossInHS_Poly = "No Solving"
 Exit Function
End If
seekCrossInHS_Poly = H_PS(P, S)
Exit Function
errLab:
    seekCrossInHS_Poly = "No Solving"
End Function


Function seekCrossInHS_Poly2(P As Double, ExpLnPoly As Variant) As Variant  'H-S图上寻找膨胀线(多项式）与等压线交点
Dim delta As Double, i As Integer, n As Integer, K As Integer
Dim S1 As Double, S2 As Double, S As Double
Dim XS() As Double
Dim Y1 As Double, Y2 As Double, y As Double

On Error GoTo errLab

If Not IsArray(ExpLnPoly) Then GoTo errLab
n = UBound(ExpLnPoly)
ReDim XS(1 To n)
delta = 0.00001
S1 = 7.2
S2 = 7.4
For i = 1 To n
  XS(i) = S1 ^ (n - 1 - i)
Next
Y1 = Application.WorksheetFunction.SumProduct(ExpLnPoly, XS) - H_PS(P, S1)

For i = 1 To n
  XS(i) = S2 ^ (n - 1 - i)
Next
Y2 = Application.WorksheetFunction.SumProduct(ExpLnPoly, XS) - H_PS(P, S2)
S = S1 - Y1 * (S1 - S2) / (Y1 - Y2)
For i = 1 To n
  XS(i) = S ^ (n - 1 - i)
Next
y = Application.WorksheetFunction.SumProduct(ExpLnPoly, XS) - H_PS(P, S)
K = 0
Do While (Abs(y) > delta And K < 50)
   K = K + 1
   S2 = S1
   Y2 = Y1
   S1 = S
   Y1 = y
  
 S = S1 - Y1 * (S1 - S2) / (Y1 - Y2)
    For i = 1 To n
      XS(i) = S ^ (n - 1 - i)
    Next
  y = Application.WorksheetFunction.SumProduct(ExpLnPoly, XS) - H_PS(P, S)
Loop
If Not K < 50 Then GoTo errLab '无解
seekCrossInHS2 = H_PS(P, S)
Exit Function
errLab:
    seekCrossInHS2 = "Error"

End Function
    
    
    
    
    
    
    
    
    
    
    
    
