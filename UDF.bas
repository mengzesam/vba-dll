Attribute VB_Name = "UDF"
Public Function mySeq(n As Integer) As Variant
Dim i As Integer, seq As Variant
ReDim seq(0 To n)
For i = 0 To n
seq(i) = i
Next
mySeq = seq
End Function

Public Function myRevSeq(n As Integer) As Variant
Dim i As Integer, seq As Variant
ReDim seq(0 To n)
For i = 0 To n
seq(i) = n - i
Next
myRevSeq = seq
End Function


