Attribute VB_Name = "vecCalc"
Option Explicit
Option Base 1

Type Complex_Double
  real As Double
  img As Double
End Type


Public Function VecAdd(ByVal v1 As String, ByVal v2 As String, Optional decim = 10) As String 'amp∠angle或amp∠angle°形式的相加
    Dim vv1 As Complex_Double, vv2 As Complex_Double, result As Complex_Double
    Dim amp As Double, angle As Double
    
    vv1 = complexConvert(v1)
    vv2 = complexConvert(v2)
    result = comAdd(vv1, vv2)
    
    amp = (result.real ^ 2 + result.img ^ 2) ^ 0.5
    angle = ThisWorkbook.Application.WorksheetFunction.Atan2(result.real, result.img) * 180# / PI
    If (angle < 0) Then angle = angle + 360
    VecAdd = amp & "∠" & angle & "°"
End Function

Public Function VecSub(ByVal v1 As String, ByVal v2 As String, Optional decim = 10) As String 'amp∠angle或amp∠angle°形式的相减
    Dim vv1 As Complex_Double, vv2 As Complex_Double, result As Complex_Double
    Dim amp As Double, angle As Double
    
    vv1 = complexConvert(v1)
    vv2 = complexConvert(v2)
    result = comSub(vv1, vv2)
    
    amp = (result.real ^ 2 + result.img ^ 2) ^ 0.5
    angle = ThisWorkbook.Application.WorksheetFunction.Atan2(result.real, result.img) * 180# / PI
    If (angle < 0) Then angle = angle + 360
    VecSub = amp & "∠" & angle & "°"
End Function

Public Function VecProduct(ByVal v1 As String, ByVal v2 As String) 'amp∠angle或amp∠angle°形式的相乘
    Dim vals1, vals2
    Dim real As Double, angle As Double
    On Error GoTo errFlag
    vals1 = Split(Replace(v1, "°", ""), "∠")
    vals2 = Split(Replace(v2, "°", ""), "∠")
    real = CDbl(vals1(0)) * CDbl(vals2(0))
    angle = CDbl(vals1(1)) + CDbl(vals2(1))
    angle = angle - Int(angle / 360) * 360
    If angle < 0 Then angle = angle + 360
    VecProduct = real & "∠" & angle & "°"
    Exit Function
errFlag:
     VecProduct = "input error"
End Function


Public Function VecDiv(ByVal v1 As String, ByVal v2 As String) 'amp∠angle或amp∠angle°形式的相除
    Dim vals1, vals2
    Dim real As Double, angle As Double
    On Error GoTo errFlag
    vals1 = Split(Replace(v1, "°", ""), "∠")
    vals2 = Split(Replace(v2, "°", ""), "∠")
    real = CDbl(vals1(0)) / CDbl(vals2(0))
    angle = CDbl(vals1(1)) - CDbl(vals2(1))
    angle = angle - Int(angle / 360) * 360
    If angle < 0 Then angle = angle + 360
    VecDiv = real & "∠" & angle & ""
    Exit Function
errFlag:
     VecDiv = "input error"
End Function


Public Function VecNegative(v1 As String) 'amp∠angle或amp∠angle°形式的负值
    Dim vals
    Dim angle As Double
    On Error GoTo errFlag
    vals = Split(Replace(v1, "°", ""), "∠")
    angle = CDbl(vals(1)) + 180
    angle = angle - Int(angle / 360) * 360
    If angle < 0 Then angle = angle + 360
    VecNegative = vals(0) & "∠" & angle & "°"
    Exit Function
errFlag:
     VecNegative = "input error"
End Function


Public Function VecAverage(myrg As Range)
  Dim real() As Double, img() As Double
  Dim amp As Double, angle As Double, avgreal As Double, avgimg As Double
  Dim size As Integer, cnt As Integer, i As Integer
  Dim rr, vals
  On Error GoTo errFlag
  amp = 0#
  cnt = 0
  size = myrg.Columns.Count * myrg.Rows.Count
  ReDim real(1 To size)
  ReDim img(1 To size)
  
  For Each rr In myrg
    vals = Split(Replace(rr.Value, "°", ""), "∠")
    amp = CDbl(vals(0))
    angle = CDbl(vals(1)) * PI / 180
    cnt = cnt + 1
    real(cnt) = amp * Cos(angle)
    img(cnt) = amp * Sin(angle)
  Next
  avgreal = 0#
  avgimg = 0#
  For i = 1 To cnt
    avgreal = avgreal + real(i)
    avgimg = avgimg + img(i)
  Next
  avgreal = avgreal / cnt
  avgimg = avgimg / cnt
  amp = (avgreal ^ 2 + avgimg ^ 2) ^ 0.5
  angle = ThisWorkbook.Application.WorksheetFunction.Atan2(avgreal, avgimg) * 180 / PI
  If (angle < 0) Then angle = angle + 360#
  VecAverage = amp & "∠" & angle & "°"
  Exit Function
errFlag:
  VecAverage = "input error"
End Function


Public Function VecDisp(v As String, Optional decim = 3)
    Dim vals
    Dim angle As Double
    On Error GoTo errFlag
    vals = Split(Replace(v, "°", ""), "∠")
    angle = CDbl(vals(1))
    angle = angle - Int(angle / 360) * 360
    If (angle < 0) Then angle = angle + 360
    VecDisp = Round(vals(0), decim) & "∠" & Round(angle, decim) & "°"
    Exit Function
errFlag:
    VecDisp = "inputError"
End Function


Private Function complexConvert(ByVal arg As String, Optional ByVal delimiter As String = "∠") As Complex_Double
    Dim vals
    Dim amp As Double, angle As Double
    Dim ans As Complex_Double
    On Error GoTo errFlag
    vals = Split(Replace(arg, "°", ""), delimiter)
    amp = CDbl(vals(0))
    angle = CDbl(vals(1)) / 180# * PI
    ans.real = amp * Cos(angle)
    ans.img = amp * Sin(angle)
    complexConvert = ans
    Exit Function
errFlag:
    ans.real = -9999999999#
    ans.img = -9999999999#
    complexConvert = ans
End Function

Private Function comAdd(v1 As Complex_Double, v2 As Complex_Double) As Complex_Double
    Dim ans As Complex_Double
    ans.real = v1.real + v2.real
    ans.img = v1.img + v2.img
    comAdd = ans
End Function

Private Function comSub(v1 As Complex_Double, v2 As Complex_Double) As Complex_Double
    Dim ans As Complex_Double
    ans.real = v1.real - v2.real
    ans.img = v1.img - v2.img
    comSub = ans
End Function

