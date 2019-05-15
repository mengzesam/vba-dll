Attribute VB_Name = "LAPACK"
Option Explicit
Option Base 1

Rem *****************************************************************************
Rem 定义复数类型*
Type lapack_complex_float
  real As Single
  img As Single
End Type

Type lapack_complex_double
  real As Double
  img As Double
End Type


Rem ******************************************************************************
Rem 定义LAPACK接口函数


Public Const LAPACK_ROW_MAJOR As Long = 101
Public Const LAPACK_COL_MAJOR As Long = 102
Public Const PI As Double = 3.14159265358979

Rem 以下函数是定义lapacke.dll的接口函数

Private Declare Function lapacke_dgesv Lib "liblapacke.dll" Alias "LAPACKE_dgesv@32" (ByVal matrix_order As Long, ByVal n As Long, ByVal nrhs As Long, _
ByRef A0 As Double, ByVal lda As Long, ByRef ipiv0 As Long, ByRef B0 As Double, ByVal ldb As Long) As Long

Private Declare Function lapacke_zgesv Lib "liblapacke.dll" Alias "LAPACKE_zgesv@32" (ByVal matrix_order As Long, ByVal n As Long, ByVal nrhs As Long, _
ByRef A0 As lapack_complex_double, ByVal lda As Long, ByRef ipiv0 As Long, ByRef B0 As lapack_complex_double, ByVal ldb As Long) As Long

Rem=======================================================================================

Rem 以下函数是dll接口函数在Excel里的实现

Public Function lapack_dgesv(A As Variant, B As Variant)
Dim matrix_order As Long
Dim n As Long
Dim m As Long
Dim nB As Long
Dim nrhs As Long
Dim lda As Long
Dim ldb As Long
Dim ipiv() As Long
Dim AA() As Double
Dim BB() As Double
Dim info As Long
Dim i As Long, j As Long
Dim X
On Error GoTo errFlag
If (TypeName(A) = "Range") Then
 n = A.Rows.Count
 m = A.Columns.Count
 If n <> m Then
   lapack_dgesv = "n<>m"
   Exit Function
 End If
 ReDim AA(1 To n * m) As Double
 For j = 1 To m
    For i = 1 To n
      AA((j - 1) * n + i) = A.Cells(i, j).Value
    Next
 Next
ElseIf (TypeName(A) = "Variant()") Then
Else
  lapack_dgesv = "A input error"
End If

If (TypeName(B) = "Range") Then
 nB = B.Rows.Count
 nrhs = B.Columns.Count
 If nB <> n Then
   lapack_dgesv = "nA<>nB"
   Exit Function
 End If
 ReDim BB(1 To nB * nrhs) As Double
 For j = 1 To nrhs
    For i = 1 To nB
      BB((j - 1) * nB + i) = B.Cells(i, j).Value
    Next
 Next
ElseIf (TypeName(B) = "Variant()") Then
Else
  lapack_dgesv = "B input error"
End If
ReDim ipiv(1 To nB * nrhs) As Long
matrix_order = LAPACK_COL_MAJOR
lda = n
ldb = nB
info = lapacke_dgesv(matrix_order, n, nrhs, AA(1), lda, ipiv(1), BB(1), ldb)
If (info <> 0) Then
  lapack_dgesv = "Lapack Err"
  Exit Function
End If
If (nrhs = 1) Then
  lapack_dgesv = Application.WorksheetFunction.Transpose(BB)
ElseIf (nrhs > 1) Then
  ReDim X(1 To n, 1 To nrhs) As Double
  For j = 1 To nrhs
   For i = 1 To n
    X(i, j) = BB((j - 1) * n + i)
   Next
  Next
  lapack_dgesv = X
End If
Exit Function
errFlag:
 lapack_dgesv = "input Err"
End Function




Public Function lapack_zgesv(A As Variant, B As Variant, Optional complex_type As Boolean = False) 'complex_type=false 按普通形式c=real+iimg
                                                                                               'complex_type=true 按极坐标形式c=abs@angle
Dim matrix_order As Long
Dim n As Long
Dim m As Long
Dim nB As Long
Dim nrhs As Long
Dim lda As Long
Dim ldb As Long
Dim ipiv() As Long
Dim AA() As lapack_complex_double
Dim BB() As lapack_complex_double
Dim info As Long
Dim i As Long, j As Long
Dim c1 As Double
Dim c2 As Double
Dim delimiter As String
Dim X
On Error GoTo errFlag
If (complex_type) Then '复数按极坐标形式
  delimiter = "∠" '"@"
Else '复数按普通标形式
  delimiter = "+i"
End If

If (TypeName(A) = "Range") Then
 n = A.Rows.Count
 m = A.Columns.Count
 If n <> m Then
   lapack_zgesv = "n<>m"
   Exit Function
 End If
 ReDim AA(1 To n * m) As lapack_complex_double
    For j = 1 To m
       For i = 1 To n
         c1 = Split(A.Cells(i, j).Value, delimiter)(0)
         c2 = ThisWorkbook.Application.WorksheetFunction.Substitute(Split(A.Cells(i, j).Value, delimiter)(1), "°", "")
         If (complex_type) Then
            AA((j - 1) * n + i).real = c1 * Cos(c2 / 180 * PI)
            AA((j - 1) * n + i).img = c1 * Sin(c2 / 180 * PI)
         Else
            AA((j - 1) * n + i).real = c1
            AA((j - 1) * n + i).img = c2
         End If
       Next
    Next
ElseIf (TypeName(A) = "Variant()") Then
Else
  lapack_zgesv = "A input error"
End If

If (TypeName(B) = "Range") Then
 nB = B.Rows.Count
 nrhs = B.Columns.Count
 If nB <> n Then
   lapack_zgesv = "nA<>nB"
   Exit Function
 End If
 ReDim BB(1 To nB * nrhs) As lapack_complex_double
   For j = 1 To nrhs
       For i = 1 To n
         c1 = Split(B.Cells(i, j).Value, delimiter)(0)
         c2 = ThisWorkbook.Application.WorksheetFunction.Substitute(Split(B.Cells(i, j).Value, delimiter)(1), "°", "")
         If (complex_type) Then
            BB((j - 1) * n + i).real = c1 * Cos(c2 / 180 * PI)
            BB((j - 1) * n + i).img = c1 * Sin(c2 / 180 * PI)
         Else
            BB((j - 1) * n + i).real = c1
            BB((j - 1) * n + i).img = c2
         End If
       Next
    Next
ElseIf (TypeName(B) = "Variant()") Then
Else
  lapack_zgesv = "B input error"
End If
ReDim ipiv(1 To nB * nrhs) As Long
matrix_order = LAPACK_COL_MAJOR
lda = n
ldb = nB
info = lapacke_zgesv(matrix_order, n, nrhs, AA(1), lda, ipiv(1), BB(1), ldb)
If (info <> 0) Then
  lapack_zgesv = "Lapack Err"
  Exit Function
End If
ReDim X(1 To n, 1 To nrhs) As String
If (complex_type) Then '复数按极坐标
    For j = 1 To nrhs
       For i = 1 To n
         c1 = BB((j - 1) * n + i).real
         c2 = BB((j - 1) * n + i).img
         X(i, j) = (c1 ^ 2 + c2 ^ 2) ^ 0.5 & delimiter & Application.WorksheetFunction.ImArgument(c1 & "+" & c2 & "i") * 180 / PI
       Next
    Next
Else '复数按普通形式
    For j = 1 To nrhs
       For i = 1 To n
        X(i, j) = BB((j - 1) * n + i).real & delimiter & X(i, j) & BB((j - 1) * n + i).img
       Next
    Next
End If
lapack_zgesv = X
Exit Function
errFlag:
 lapack_zgesv = "input Err"
End Function


