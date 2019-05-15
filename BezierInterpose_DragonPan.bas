Attribute VB_Name = "BezierInterpose_DragonPan"
'   Excel的平滑线散点图，可以根据两组分别代表X-Y坐标的散点数值产生曲线图
'   但是，却没有提供这个曲线图的公式，所以无法查找曲线上的点坐标
'   后来我在以下这个网页找到了详细的说明和示例程序
'..............................................................................
'http://www.xlrotor.com/Smooth_curve_bezier_example_file.zip
'..............................................................................
'   根据其中采用的算法，进一步增添根据X坐标求Y坐标，或根据Y坐标求X坐标，更切合实际需求
'   这个自定义函数按照Excel的曲线算法(三次贝塞尔分段插值),计算平滑曲线上任意一点的点坐标
'
'   Excel的平滑曲线的大致算法是:
'   给出了两组X-Y数值以后，每一对X-Y坐标称为节点，然后在每两个节点之间画出三次贝塞尔曲线(下面简称曲线)
'   贝塞尔曲线的算法网上有很多资源，这里不介绍了，只作简单说明
'   每条曲线都由四个节点开始，计算出四个贝塞尔控制点，然后根据控制点画出唯一一条曲线
'   假设曲线的源数据是节点1,节点2,节点3,节点4(Dot1,Dot2,Dot3,Dot4)
'   那么贝塞尔控制点的计算如下                  程序作者: 海底眼(Mr. Dragon Pan)
'   Dot2是第一个控制点,也是曲点的起点，Dot3是第四个控制点也是曲线的终点
'
'   第二个控制点的位置是:
'       过第一个控制点(Dot2,起点)，与Dot1, Dot3的连线平行，且与Dot2距离为 1/6 * 线段Dot1_Dot3的长度
'       假如是图形的第一段曲线，取节点1,1,2,3进行计算,即 Dot2 = Dot1
'       且第二个控制点与第一控制点距离取 1/3 * |Dot1_Dot3|,而不是1/6 * |Dot1_Dot3|
'       假如 1/2 * |Dot2_Dot3| < 1/6 * |Dot1_Dot3|
'       那么第二个控制点与第一控制点距离取  1/2 * |Dot2_Dot3|,而不是1/6 * |Dot1_Dot3|
'
'   第三个控制点的位置是:
'       过第四个控制点(Dot3,终点)，与Dot2, Dot4的连线平行，且与Dot3距离为 1/6 * |Dot2_Dot4|
'       假如是图形的最后一段曲线，取节点Last-2,Last-1,Last,Last进行计算,即 Dot4 = Dot3
'       且第三个控制点与第四控制点距离取 1/3 * |Dot2_Dot4|,而不是1/6 * |Dot2_Dot4|
'       假如 1/2 * |Dot2_Dot3| < 1/6 * |Dot2_Dot4|
'       那么第二个控制点与第一控制点距离取  1/2 * |Dot2_Dot4|,而不是1/6 * |Dot2_Dot4|
'...............................................................................................
'   这个自定义函数的计算流程是
'   Step1: 检查输入的X-Y数值是否有错误，如(输入不够三个点，X-Y的数量不一致,起始搜索节点超过范围等等)
'   Step2: 从参数指定的节点开始，计算出四个贝塞尔控制点，得到贝塞尔插值多项式方程，
'          然后代入已知的待求数值，看它能不能满足 f(t)=0 有解 (即曲线包含待查数值)
'   Step3: 如果 f(t)=0 有解，根据解出来的 t 值计算X-Y坐标，退出程序,否则继续检查下一段曲线
'   Step4: 如果所有分段曲线都不包含待查数值，退出程序
'...............................................................................................

Option Base 1       '所有数组的第一个元素编号为1(默认为0)
Type Vector         '自定义数据结构(用二维向量代表坐标系里面的点坐标)
    X As Double
    y As Double
End Type

Const NoError = "No error"      '错误提示信息
Const Error1 = "Error: The size of known_x must equal to size of known_y"
Const Error2 = "Error: The size of known_x must equal to or greater than 3"
Const Error3 = "Error: StartKnot must be >=1 and <=count(known_x)-1"
Const Error4 = "Error: known_value_type must be ""x"",""y"",or ""t"" "
Const Error5 = "Error: When known_value_type is ""t"" , known_value must >=0 and <=1"
Const Error10 = "Error: known_value is not on the curve (defined by given known_x and known_y)"
Const NoRoot = "No Root"
Const MaxErr = 0.00000001
Const MaxLoop = 1000

Dim SizeX, SizeY As Long        '输入区域的大小
Dim Dot1 As Vector              '输入区域里面，用作计算贝塞尔控制点的四个节点
Dim Dot2 As Vector
Dim Dot3 As Vector
Dim Dot4 As Vector
Dim BezierPt1 As Vector         '生成贝塞尔曲线的四个贝塞尔控制点
Dim BezierPt2 As Vector
Dim BezierPt3 As Vector
Dim BezierPt4 As Vector
Dim OffsetTo2 As Vector         '第二,三个贝塞尔控制点跟起点，终点的距离关系
Dim OffsetTo3 As Vector

Dim ValueType As Variant        '输入待查数值的类型,"x"代表输入的是X坐标，求对应的Y坐标
Dim Interpol_here As Boolean    '当前分段曲线是否包含待查数值
Dim key_value, a, b, c, d As Double     '贝塞尔曲线插值多项式的系数
Dim t1, t2, t3 As Variant               '贝塞尔曲线插值多项式的根
Dim a3, a2, a1, a0 As Double
'-------------------------------------------------------------------------------------------------
'主程序开始，至少要输入三个参数，第一个是X坐标系列，然后是Y坐标系列，第三个是待查数值
'第四个参数是从哪一段曲线开始查找，如果曲线可以返回多个值，那么分别指定起始节点就可以找出全部合要求的点
'第五个参数是待查数值的类型,"x"代表输入x坐标求对应y坐标，"y"则相反,"t"是直接输入贝塞尔插值多项式的参数
'-------------------------------------------------------------------------------------------------
Function BezierFit(Known_X, Known_Y As Range, known_value, Optional StartKnot As Long = 1, Optional known_value_type As Variant = "x") As Variant
Dim j As Long
Dim x1Value, y1Value, x2Value, y2Value, x3Value, y3Value As Variant
Dim ErrorMsg As Variant

ValueType = LCase(known_value_type)     '待查数值的类型转化为小写，并赋值到全局变量ValueType
key_value = known_value                 '待查数值赋值到全局变量key_value

ErrorMsg = ErrorCheck(Known_X, Known_Y, StartKnot)  '检查输入错误
If ErrorMsg <> NoError Then                         '有错误就返回错误信息，退出程序
    BezierFit = Array(ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg)
    Exit Function
End If

For j = StartKnot To SizeX - 1              '从指定的节点开始，没有指定节点就从1开始
    Call FindFourDots(Known_X, Known_Y, j)  '找出输入X-Y点坐标里面，应该用于计算的四个结点
    Call FindFourBezierPoints(Dot1, Dot2, Dot3, Dot4)   '根据四个结点计算四个贝塞尔控制点
    Call FindABCD                           '根据待查数值的类型，和贝塞尔控制点,计算贝塞尔插值多项式的系数
    Call Find_t                             '检查贝塞尔曲线是否包含待查数值
    If Interpol_here = True Then Exit For
Next j

If Interpol_here = True Then    '计算点坐标，并返回
                                '以下是由四个贝塞尔控制点决定的,贝塞尔曲线的参数方程
    x1Value = (1 - t1) ^ 3 * BezierPt1.X + 3 * t1 * (1 - t1) ^ 2 * BezierPt2.X + 3 * t1 ^ 2 * (1 - t1) * BezierPt3.X + t1 ^ 3 * BezierPt4.X
    y1Value = (1 - t1) ^ 3 * BezierPt1.y + 3 * t1 * (1 - t1) ^ 2 * BezierPt2.y + 3 * t1 ^ 2 * (1 - t1) * BezierPt3.y + t1 ^ 3 * BezierPt4.y
    x2Value = (1 - t2) ^ 3 * BezierPt1.X + 3 * t2 * (1 - t2) ^ 2 * BezierPt2.X + 3 * t2 ^ 2 * (1 - t2) * BezierPt3.X + t2 ^ 3 * BezierPt4.X
    y2Value = (1 - t2) ^ 3 * BezierPt1.y + 3 * t2 * (1 - t2) ^ 2 * BezierPt2.y + 3 * t2 ^ 2 * (1 - t2) * BezierPt3.y + t2 ^ 3 * BezierPt4.y
    x3Value = (1 - t3) ^ 3 * BezierPt1.X + 3 * t3 * (1 - t3) ^ 2 * BezierPt2.X + 3 * t3 ^ 2 * (1 - t3) * BezierPt3.X + t3 ^ 3 * BezierPt4.X
    y3Value = (1 - t3) ^ 3 * BezierPt1.y + 3 * t3 * (1 - t3) ^ 2 * BezierPt2.y + 3 * t3 ^ 2 * (1 - t3) * BezierPt3.y + t3 ^ 3 * BezierPt4.y
    BezierFit = Array(x1Value, y1Value, x2Value, y2Value, x3Value, y3Value)
Else
    BezierFit = Array(Error10, Error10, Error10, Error10, Error10, Error10)
End If

End Function

Function ErrorCheck(Known_X, Known_Y, StartKnot) As Variant

ErrorCheck = NoError
SizeX = Known_X.Count
SizeY = Known_Y.Count

If SizeX <> SizeY Then  '假如输入的X坐标数目不等于Y坐标数目
ErrorCheck = Error1
Exit Function
End If

If SizeX < 3 Then       '输入的X-Y坐标对少于三个
ErrorCheck = Error2
Exit Function
End If

If (StartKnot < 1 Or StartKnot >= SizeX) Then   '指定的起始节点超出范围
ErrorCheck = Error3
Exit Function
End If

If (ValueType <> "x" And ValueType <> "y" And ValueType <> "t") Then    '输入的待查数值类型不是x, y, t
ErrorCheck = Error4
Exit Function
End If

If ((ValueType = "t" And key_value > 1) Or (ValueType = "t" And keyvalue < 0)) Then     ' t 类型的范围是0-1
ErrorCheck = Error5
Exit Function
End If

End Function

Sub FindFourDots(Known_X, Known_Y, j)   '根据X-Y数值，及起始节点，找出用于计算的四个结点坐标
    If j = 1 Then                       '第一个结点 Dot2 = Dot1
        Dot1.X = Known_X(1)
        Dot1.y = Known_Y(1)
     Else
        Dot1.X = Known_X(j - 1)
        Dot1.y = Known_Y(j - 1)
     End If
     
     Dot2.X = Known_X(j)
     Dot2.y = Known_Y(j)
     Dot3.X = Known_X(j + 1)
     Dot3.y = Known_Y(j + 1)
     
     If j = SizeX - 1 Then              '最后一个结点 Dot4 = Dot3
        Dot4.X = Dot3.X
        Dot4.y = Dot3.y
     Else
        Dot4.X = Known_X(j + 2)
        Dot4.y = Known_Y(j + 2)
     End If
End Sub

Sub FindFourBezierPoints(Dot1 As Vector, Dot2 As Vector, Dot3 As Vector, Dot4 As Vector)
Dim d12, d23, d34, d13, d14, d24 As Double
d12 = DistAtoB(Dot1, Dot2)      '计算平面坐标系上的两点距离
d23 = DistAtoB(Dot2, Dot3)
d34 = DistAtoB(Dot3, Dot4)
d13 = DistAtoB(Dot1, Dot3)
d14 = DistAtoB(Dot1, Dot4)
d24 = DistAtoB(Dot2, Dot4)

BezierPt1 = Dot2
BezierPt4 = Dot3
OffsetTo2 = AsubB(Dot3, Dot1)   '向量减法
OffsetTo3 = AsubB(Dot2, Dot4)

If ((d13 / 6 < d23 / 2) And (d24 / 6 < d23 / 2)) Then
    If (Dot1.X <> Dot2.X Or Dot1.y <> Dot2.y) Then OffsetTo2 = AmultiF(OffsetTo2, 1 / 6)
    If (Dot1.X = Dot2.X And Dot1.y = Dot2.y) Then OffsetTo2 = AmultiF(OffsetTo2, 1 / 3)
    If (Dot3.X <> Dot4.X Or Dot3.y <> Dot4.y) Then OffsetTo3 = AmultiF(OffsetTo3, 1 / 6)
    If (Dot3.X = Dot4.X And Dot3.y = Dot4.y) Then OffsetTo3 = AmultiF(OffsetTo3, 1 / 3)
ElseIf ((d13 / 6 >= d23 / 2) And (d24 / 6 >= d23 / 2)) Then
    OffsetTo2 = AmultiF(OffsetTo2, d23 / 12)
    OffsetTo3 = AmultiF(OffsetTo3, d23 / 12)
ElseIf (d13 / 6 >= d23 / 2) Then
    OffsetTo2 = AmultiF(OffsetTo2, d23 / 2 / d13)
    OffsetTo3 = AmultiF(OffsetTo3, d23 / 2 / d13)
ElseIf (d24 / 6 >= d23 / 2) Then
    OffsetTo2 = AmultiF(OffsetTo2, d23 / 2 / d24)
    OffsetTo3 = AmultiF(OffsetTo3, d23 / 2 / d24)
End If

BezierPt2 = AaddB(BezierPt1, OffsetTo2)     '向量加法
BezierPt3 = AaddB(BezierPt4, OffsetTo3)

End Sub
Function DistAtoB(dota As Vector, dotb As Vector) As Double
DistAtoB = ((dota.X - dotb.X) ^ 2 + (dota.y - dotb.y) ^ 2) ^ 0.5
End Function
Function AaddB(dota As Vector, dotb As Vector) As Vector
AaddB.X = dota.X + dotb.X
AaddB.y = dota.y + dotb.y
End Function
Function AsubB(dota As Vector, dotb As Vector) As Vector
AsubB.X = dota.X - dotb.X
AsubB.y = dota.y - dotb.y
End Function
Function AmultiF(dota As Vector, MultiFactor As Double) As Vector
AmultiF.X = dota.X * MultiFactor
AmultiF.y = dota.y * MultiFactor
End Function

Sub FindABCD()

If ValueType = "x" Then     '参数类型是x, 需要解参数方程 f(t) = x，这里设定参数方程的系数
a = -BezierPt1.X + 3 * BezierPt2.X - 3 * BezierPt3.X + BezierPt4.X
b = 3 * BezierPt1.X - 6 * BezierPt2.X + 3 * BezierPt3.X
c = -3 * BezierPt1.X + 3 * BezierPt2.X
d = BezierPt1.X - key_value
End If

If ValueType = "y" Then    '参数类型是x, 需要解参数方程 f(t) = y，这里设定参数方程的系数
a = -BezierPt1.y + 3 * BezierPt2.y - 3 * BezierPt3.y + BezierPt4.y
b = 3 * BezierPt1.y - 6 * BezierPt2.y + 3 * BezierPt3.y
c = -3 * BezierPt1.y + 3 * BezierPt2.y
d = BezierPt1.y - key_value
End If
End Sub

Sub Find_t()        '计算当 f(t) = 待查数值时, t应该是什么数值

Dim tArr As Variant

Interpol_here = True

If ValueType = "t" Then     '待查数值类型为t,那么无需计算
    t1 = key_value
    t2 = key_value
    t3 = key_value
    Exit Sub
End If

tArr = Solve_Order3_Equation(a, b, c, d)    '否则，解三次贝塞尔参数方程 f(t) = 待查数值
t1 = tArr(1)                                '解得方程的三个根
t2 = tArr(2)
t3 = tArr(3)

If (t1 > 1 Or t1 < 0) Then                  '参数方程的 t 值范围应该是 0-1
    t1 = NoRoot
End If
If (t2 > 1 Or t2 < 0) Then
    t2 = NoRoot
End If
If (t3 > 1 Or t3 < 0) Then
    t3 = NoRoot
End If

If (IsNumeric(t1) = False And IsNumeric(t2) = False And IsNumeric(t3) = False) Then
    Interpol_here = False
End If                      '   三个根都不合要求，代表曲线上没有包含待查数值的点

If (t1 = NoRoot And t2 <> NoRoot) Then  '至少有一个根，则用它代替NoRoot的结果,方便Excel画图
    t1 = t2
End If
If (t1 = NoRoot And t3 <> NoRoot) Then
    t1 = t3
End If

If (t2 = NoRoot) Then t2 = t1
If (t3 = NoRoot) Then t3 = t1

End Sub

'................................................................................................
'   牛顿法解三次方程，先求解方程的导函数，得到方程的拐点(导数等于0的点)
'   然后分三段用迭代法分别求三个根
'................................................................................................
Public Function Solve_Order3_Equation(p3, p2, p1, P0, Optional Starting As Double = -10000000000#, Optional Ending As Double = 10000000000#) As Variant
Dim Two_X, TurningPoint, x1, x2, x3 As Variant
Dim X As Double
a3 = p3
a2 = p2
a1 = p1
a0 = P0
x1 = NoRoot
x2 = NoRoot
x3 = NoRoot

x1 = Newton_Solve(Starting)
If a3 = 0 Then                                  '   如果三次方程没有三次项
    Two_X = Solve_Order2_Equation(a2, a1, a0)   '   解释法直接求二次方程的解
    x1 = Two_X(1)
    x2 = Two_X(2)
Else
    TurningPoint = Solve_Order2_Equation(3 * a3, 2 * a2, 1 * a1)    '   求解 f'(t) = 0
    
If (TurningPoint(1) = NoRoot And TurningPoint(2) = NoRoot) Then     '   分段求根
        X = 0
        x1 = Newton_Solve(X)
    ElseIf (TurningPoint(1) <> NoRoot And TurningPoint(2) = NoRoot) Then
        If f_x(Starting) * f_x(TurningPoint(1)) < 0 Then
            X = (Starting + TurningPoint(1)) / 2
            x1 = Newton_Solve(X)
        End If
        If f_x(TurningPoint(2)) * f_x(Ending) < 0 Then
            X = (TurningPoint(2) + Ending) / 2
            x3 = Newton_Solve(X)
        End If
    ElseIf (TurningPoint(1) <> NoRoot And TurningPoint(2) <> NoRoot) Then
        If f_x(Starting) * f_x(TurningPoint(1)) < 0 Then
            X = (Starting + TurningPoint(1)) / 2
            x1 = Newton_Solve(X)
        End If
        If f_x(TurningPoint(1)) * f_x(TurningPoint(2)) < 0 Then
            X = (TurningPoint(1) + TurningPoint(2)) / 2
            x2 = Newton_Solve(X)
        End If
        If f_x(TurningPoint(2)) * f_x(Ending) < 0 Then
            X = (TurningPoint(2) + Ending) / 2
            x3 = Newton_Solve(X)
        End If
    End If
End If

Solve_Order3_Equation = Array(x1, x2, x3)

End Function

Function f_x(xValue) As Double                  ' f_x = f(x) 求贝塞尔参数方程 f(t)的值
f_x = a3 * xValue ^ 3 + a2 * xValue ^ 2 + a1 * xValue + a0
End Function
Function Df_x(xValue As Double) As Double       ' Df_x = f'(x) ' f_x = f(x) 求贝塞尔参数方程导函数 f'(t)的值
Df_x = 3 * a3 * xValue ^ 2 + 2 * a2 * xValue + a1
End Function

Function Solve_Order2_Equation(k2, k1, k0 As Double) As Variant
Dim b2SUB4ac As Double

If (k2 = 0) Then
    If k1 = 0 Then
            Solve_Order2_Equation = Array(NoRoot, NoRoot)
            Exit Function
        ElseIf (k1 <> 0 And k0 = 0) Then
            Solve_Order2_Equation = Array(0, 0)
            Exit Function
        ElseIf (k1 <> 0 And k0 <> 0) Then
            Solve_Order2_Equation = Array(-k0 / k1, -k0 / k1)
            Exit Function
    End If
End If

b2SUB4ac = (k1) ^ 2 - 4 * k2 * k0       ' 二次方程可以直接用公式求解,b^2-4*a*c
If b2SUB4ac < 0 Then
    Solve_Order2_Equation = Array(NoRoot, NoRoot)
End If

If b2SUB4ac = 0 Then
    Solve_Order2_Equation = Array(-k1 / 2 / k2, -k1 / 2 / k2)
End If

If b2SUB4ac > 0 Then
    If (-k1 + b2SUB4ac ^ 0.5) / 2 / k2 < (-k1 - b2SUB4ac ^ 0.5) / 2 / k2 Then
        Solve_Order2_Equation = Array((-k1 + b2SUB4ac ^ 0.5) / 2 / k2, (-k1 - b2SUB4ac ^ 0.5) / 2 / k2)
    Else
        Solve_Order2_Equation = Array((-k1 - b2SUB4ac ^ 0.5) / 2 / k2, (-k1 + b2SUB4ac ^ 0.5) / 2 / k2)
    End If
End If

End Function

Function Newton_Solve(x0 As Double) As Variant
Dim i, eps As Double

i = 0
eps = Abs(f_x(x0))
Do While (eps > MaxErr)                     '如果取初值，函数的绝对值大于允许误差
    If (Df_x(x0) <> 0 And i < MaxLoop) Then '而且分子不等于０，没有超出最大迭代次数
        x0 = x0 - f_x(x0) / Df_x(x0)        '牛顿法求下一个值 x' = x0 - f(x0) / f'(x0)                          程序作者: 海底眼(Mr. Dragon Pan)
        eps = Abs(f_x(x0))
        i = i + 1
    Else
        Newton_Solve = NoRoot
        Exit Function
    End If
Loop

Newton_Solve = x0
End Function


