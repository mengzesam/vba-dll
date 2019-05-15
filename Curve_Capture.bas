Attribute VB_Name = "Curve_Capture"
Option Explicit
Type MYPOINT
  X As Long
  y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (ByRef curPoint As MYPOINT) As Long

