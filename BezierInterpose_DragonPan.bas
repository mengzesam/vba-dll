Attribute VB_Name = "BezierInterpose_DragonPan"
'   Excel��ƽ����ɢ��ͼ�����Ը�������ֱ����X-Y�����ɢ����ֵ��������ͼ
'   ���ǣ�ȴû���ṩ�������ͼ�Ĺ�ʽ�������޷����������ϵĵ�����
'   �����������������ҳ�ҵ�����ϸ��˵����ʾ������
'..............................................................................
'http://www.xlrotor.com/Smooth_curve_bezier_example_file.zip
'..............................................................................
'   �������в��õ��㷨����һ���������X������Y���꣬�����Y������X���꣬���к�ʵ������
'   ����Զ��庯������Excel�������㷨(���α������ֶβ�ֵ),����ƽ������������һ��ĵ�����
'
'   Excel��ƽ�����ߵĴ����㷨��:
'   ����������X-Y��ֵ�Ժ�ÿһ��X-Y�����Ϊ�ڵ㣬Ȼ����ÿ�����ڵ�֮�仭�����α���������(����������)
'   ���������ߵ��㷨�����кܶ���Դ�����ﲻ�����ˣ�ֻ����˵��
'   ÿ�����߶����ĸ��ڵ㿪ʼ��������ĸ����������Ƶ㣬Ȼ����ݿ��Ƶ㻭��Ψһһ������
'   �������ߵ�Դ�����ǽڵ�1,�ڵ�2,�ڵ�3,�ڵ�4(Dot1,Dot2,Dot3,Dot4)
'   ��ô���������Ƶ�ļ�������                  ��������: ������(Mr. Dragon Pan)
'   Dot2�ǵ�һ�����Ƶ�,Ҳ���������㣬Dot3�ǵ��ĸ����Ƶ�Ҳ�����ߵ��յ�
'
'   �ڶ������Ƶ��λ����:
'       ����һ�����Ƶ�(Dot2,���)����Dot1, Dot3������ƽ�У�����Dot2����Ϊ 1/6 * �߶�Dot1_Dot3�ĳ���
'       ������ͼ�εĵ�һ�����ߣ�ȡ�ڵ�1,1,2,3���м���,�� Dot2 = Dot1
'       �ҵڶ������Ƶ����һ���Ƶ����ȡ 1/3 * |Dot1_Dot3|,������1/6 * |Dot1_Dot3|
'       ���� 1/2 * |Dot2_Dot3| < 1/6 * |Dot1_Dot3|
'       ��ô�ڶ������Ƶ����һ���Ƶ����ȡ  1/2 * |Dot2_Dot3|,������1/6 * |Dot1_Dot3|
'
'   ���������Ƶ��λ����:
'       �����ĸ����Ƶ�(Dot3,�յ�)����Dot2, Dot4������ƽ�У�����Dot3����Ϊ 1/6 * |Dot2_Dot4|
'       ������ͼ�ε����һ�����ߣ�ȡ�ڵ�Last-2,Last-1,Last,Last���м���,�� Dot4 = Dot3
'       �ҵ��������Ƶ�����Ŀ��Ƶ����ȡ 1/3 * |Dot2_Dot4|,������1/6 * |Dot2_Dot4|
'       ���� 1/2 * |Dot2_Dot3| < 1/6 * |Dot2_Dot4|
'       ��ô�ڶ������Ƶ����һ���Ƶ����ȡ  1/2 * |Dot2_Dot4|,������1/6 * |Dot2_Dot4|
'...............................................................................................
'   ����Զ��庯���ļ���������
'   Step1: ��������X-Y��ֵ�Ƿ��д�����(���벻�������㣬X-Y��������һ��,��ʼ�����ڵ㳬����Χ�ȵ�)
'   Step2: �Ӳ���ָ���Ľڵ㿪ʼ��������ĸ����������Ƶ㣬�õ���������ֵ����ʽ���̣�
'          Ȼ�������֪�Ĵ�����ֵ�������ܲ������� f(t)=0 �н� (�����߰���������ֵ)
'   Step3: ��� f(t)=0 �н⣬���ݽ������ t ֵ����X-Y���꣬�˳�����,������������һ������
'   Step4: ������зֶ����߶�������������ֵ���˳�����
'...............................................................................................

Option Base 1       '��������ĵ�һ��Ԫ�ر��Ϊ1(Ĭ��Ϊ0)
Type Vector         '�Զ������ݽṹ(�ö�ά������������ϵ����ĵ�����)
    X As Double
    y As Double
End Type

Const NoError = "No error"      '������ʾ��Ϣ
Const Error1 = "Error: The size of known_x must equal to size of known_y"
Const Error2 = "Error: The size of known_x must equal to or greater than 3"
Const Error3 = "Error: StartKnot must be >=1 and <=count(known_x)-1"
Const Error4 = "Error: known_value_type must be ""x"",""y"",or ""t"" "
Const Error5 = "Error: When known_value_type is ""t"" , known_value must >=0 and <=1"
Const Error10 = "Error: known_value is not on the curve (defined by given known_x and known_y)"
Const NoRoot = "No Root"
Const MaxErr = 0.00000001
Const MaxLoop = 1000

Dim SizeX, SizeY As Long        '��������Ĵ�С
Dim Dot1 As Vector              '�����������棬�������㱴�������Ƶ���ĸ��ڵ�
Dim Dot2 As Vector
Dim Dot3 As Vector
Dim Dot4 As Vector
Dim BezierPt1 As Vector         '���ɱ��������ߵ��ĸ����������Ƶ�
Dim BezierPt2 As Vector
Dim BezierPt3 As Vector
Dim BezierPt4 As Vector
Dim OffsetTo2 As Vector         '�ڶ�,�������������Ƶ����㣬�յ�ľ����ϵ
Dim OffsetTo3 As Vector

Dim ValueType As Variant        '���������ֵ������,"x"�����������X���꣬���Ӧ��Y����
Dim Interpol_here As Boolean    '��ǰ�ֶ������Ƿ����������ֵ
Dim key_value, a, b, c, d As Double     '���������߲�ֵ����ʽ��ϵ��
Dim t1, t2, t3 As Variant               '���������߲�ֵ����ʽ�ĸ�
Dim a3, a2, a1, a0 As Double
'-------------------------------------------------------------------------------------------------
'������ʼ������Ҫ����������������һ����X����ϵ�У�Ȼ����Y����ϵ�У��������Ǵ�����ֵ
'���ĸ������Ǵ���һ�����߿�ʼ���ң�������߿��Է��ض��ֵ����ô�ֱ�ָ����ʼ�ڵ�Ϳ����ҳ�ȫ����Ҫ��ĵ�
'����������Ǵ�����ֵ������,"x"��������x�������Ӧy���꣬"y"���෴,"t"��ֱ�����뱴������ֵ����ʽ�Ĳ���
'-------------------------------------------------------------------------------------------------
Function BezierFit(Known_X, Known_Y As Range, known_value, Optional StartKnot As Long = 1, Optional known_value_type As Variant = "x") As Variant
Dim j As Long
Dim x1Value, y1Value, x2Value, y2Value, x3Value, y3Value As Variant
Dim ErrorMsg As Variant

ValueType = LCase(known_value_type)     '������ֵ������ת��ΪСд������ֵ��ȫ�ֱ���ValueType
key_value = known_value                 '������ֵ��ֵ��ȫ�ֱ���key_value

ErrorMsg = ErrorCheck(Known_X, Known_Y, StartKnot)  '����������
If ErrorMsg <> NoError Then                         '�д���ͷ��ش�����Ϣ���˳�����
    BezierFit = Array(ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg, ErrorMsg)
    Exit Function
End If

For j = StartKnot To SizeX - 1              '��ָ���Ľڵ㿪ʼ��û��ָ���ڵ�ʹ�1��ʼ
    Call FindFourDots(Known_X, Known_Y, j)  '�ҳ�����X-Y���������棬Ӧ�����ڼ�����ĸ����
    Call FindFourBezierPoints(Dot1, Dot2, Dot3, Dot4)   '�����ĸ��������ĸ����������Ƶ�
    Call FindABCD                           '���ݴ�����ֵ�����ͣ��ͱ��������Ƶ�,���㱴������ֵ����ʽ��ϵ��
    Call Find_t                             '��鱴���������Ƿ����������ֵ
    If Interpol_here = True Then Exit For
Next j

If Interpol_here = True Then    '��������꣬������
                                '���������ĸ����������Ƶ������,���������ߵĲ�������
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

If SizeX <> SizeY Then  '���������X������Ŀ������Y������Ŀ
ErrorCheck = Error1
Exit Function
End If

If SizeX < 3 Then       '�����X-Y�������������
ErrorCheck = Error2
Exit Function
End If

If (StartKnot < 1 Or StartKnot >= SizeX) Then   'ָ������ʼ�ڵ㳬����Χ
ErrorCheck = Error3
Exit Function
End If

If (ValueType <> "x" And ValueType <> "y" And ValueType <> "t") Then    '����Ĵ�����ֵ���Ͳ���x, y, t
ErrorCheck = Error4
Exit Function
End If

If ((ValueType = "t" And key_value > 1) Or (ValueType = "t" And keyvalue < 0)) Then     ' t ���͵ķ�Χ��0-1
ErrorCheck = Error5
Exit Function
End If

End Function

Sub FindFourDots(Known_X, Known_Y, j)   '����X-Y��ֵ������ʼ�ڵ㣬�ҳ����ڼ�����ĸ��������
    If j = 1 Then                       '��һ����� Dot2 = Dot1
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
     
     If j = SizeX - 1 Then              '���һ����� Dot4 = Dot3
        Dot4.X = Dot3.X
        Dot4.y = Dot3.y
     Else
        Dot4.X = Known_X(j + 2)
        Dot4.y = Known_Y(j + 2)
     End If
End Sub

Sub FindFourBezierPoints(Dot1 As Vector, Dot2 As Vector, Dot3 As Vector, Dot4 As Vector)
Dim d12, d23, d34, d13, d14, d24 As Double
d12 = DistAtoB(Dot1, Dot2)      '����ƽ������ϵ�ϵ��������
d23 = DistAtoB(Dot2, Dot3)
d34 = DistAtoB(Dot3, Dot4)
d13 = DistAtoB(Dot1, Dot3)
d14 = DistAtoB(Dot1, Dot4)
d24 = DistAtoB(Dot2, Dot4)

BezierPt1 = Dot2
BezierPt4 = Dot3
OffsetTo2 = AsubB(Dot3, Dot1)   '��������
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

BezierPt2 = AaddB(BezierPt1, OffsetTo2)     '�����ӷ�
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

If ValueType = "x" Then     '����������x, ��Ҫ��������� f(t) = x�������趨�������̵�ϵ��
a = -BezierPt1.X + 3 * BezierPt2.X - 3 * BezierPt3.X + BezierPt4.X
b = 3 * BezierPt1.X - 6 * BezierPt2.X + 3 * BezierPt3.X
c = -3 * BezierPt1.X + 3 * BezierPt2.X
d = BezierPt1.X - key_value
End If

If ValueType = "y" Then    '����������x, ��Ҫ��������� f(t) = y�������趨�������̵�ϵ��
a = -BezierPt1.y + 3 * BezierPt2.y - 3 * BezierPt3.y + BezierPt4.y
b = 3 * BezierPt1.y - 6 * BezierPt2.y + 3 * BezierPt3.y
c = -3 * BezierPt1.y + 3 * BezierPt2.y
d = BezierPt1.y - key_value
End If
End Sub

Sub Find_t()        '���㵱 f(t) = ������ֵʱ, tӦ����ʲô��ֵ

Dim tArr As Variant

Interpol_here = True

If ValueType = "t" Then     '������ֵ����Ϊt,��ô�������
    t1 = key_value
    t2 = key_value
    t3 = key_value
    Exit Sub
End If

tArr = Solve_Order3_Equation(a, b, c, d)    '���򣬽����α������������� f(t) = ������ֵ
t1 = tArr(1)                                '��÷��̵�������
t2 = tArr(2)
t3 = tArr(3)

If (t1 > 1 Or t1 < 0) Then                  '�������̵� t ֵ��ΧӦ���� 0-1
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
End If                      '   ������������Ҫ�󣬴���������û�а���������ֵ�ĵ�

If (t1 = NoRoot And t2 <> NoRoot) Then  '������һ����������������NoRoot�Ľ��,����Excel��ͼ
    t1 = t2
End If
If (t1 = NoRoot And t3 <> NoRoot) Then
    t1 = t3
End If

If (t2 = NoRoot) Then t2 = t1
If (t3 = NoRoot) Then t3 = t1

End Sub

'................................................................................................
'   ţ�ٷ������η��̣�����ⷽ�̵ĵ��������õ����̵Ĺյ�(��������0�ĵ�)
'   Ȼ��������õ������ֱ���������
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
If a3 = 0 Then                                  '   ������η���û��������
    Two_X = Solve_Order2_Equation(a2, a1, a0)   '   ���ͷ�ֱ������η��̵Ľ�
    x1 = Two_X(1)
    x2 = Two_X(2)
Else
    TurningPoint = Solve_Order2_Equation(3 * a3, 2 * a2, 1 * a1)    '   ��� f'(t) = 0
    
If (TurningPoint(1) = NoRoot And TurningPoint(2) = NoRoot) Then     '   �ֶ����
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

Function f_x(xValue) As Double                  ' f_x = f(x) �������������� f(t)��ֵ
f_x = a3 * xValue ^ 3 + a2 * xValue ^ 2 + a1 * xValue + a0
End Function
Function Df_x(xValue As Double) As Double       ' Df_x = f'(x) ' f_x = f(x) �������������̵����� f'(t)��ֵ
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

b2SUB4ac = (k1) ^ 2 - 4 * k2 * k0       ' ���η��̿���ֱ���ù�ʽ���,b^2-4*a*c
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
Do While (eps > MaxErr)                     '���ȡ��ֵ�������ľ���ֵ�����������
    If (Df_x(x0) <> 0 And i < MaxLoop) Then '���ҷ��Ӳ����ڣ���û�г�������������
        x0 = x0 - f_x(x0) / Df_x(x0)        'ţ�ٷ�����һ��ֵ x' = x0 - f(x0) / f'(x0)                          ��������: ������(Mr. Dragon Pan)
        eps = Abs(f_x(x0))
        i = i + 1
    Else
        Newton_Solve = NoRoot
        Exit Function
    End If
Loop

Newton_Solve = x0
End Function


