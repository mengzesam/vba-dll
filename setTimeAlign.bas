Attribute VB_Name = "setTimeAlign"
Public Sub setTimeAlign1()
  Dim ref As Range, src As Range
  Dim desSheetName As String
  Dim Count As Long
  Dim stPos As Long, endPos As Long
  stRow = 2
  stCol = 1
  halfTimeError = 15
  Set curSht = ActiveSheet
  desSheetName = "时间对齐数据表"
  On Error Resume Next
  Set desSht = Sheets(desSheetName)
  If desSht Is Nothing Then
    Set desSht = Sheets.Add
    desSht.Name = desSheetName
  Else
    desSht.Range("A1:XFD10000000").Value = ""
  End If
  On Error GoTo exitFlag
  rowNum = Application.WorksheetFunction.CountA(Range(Chr(64 + stCol) & ":" & Chr(64 + stCol)))
  colNum = Application.WorksheetFunction.CountA(Range(stRow & ":" & stRow))
  Count = rowNum - stRow + 1
  Count2 = Application.WorksheetFunction.CountA(Range(Chr(64 + stCol + 1) & ":" & Chr(64 + stCol + 1))) - stRow + 1
  Set ref = Cells(stRow, stCol)
  Set src = Cells(stRow, stCol + 1)
  addr1 = Cells(1, stCol).Address & ":" & Cells(rowNum + stRow, stCol).Address
  desSht.Range(addr1).Value = Range(addr1).Value
  addr1 = Cells(1, stCol).Address & ":" & Cells(stRow - 1, colNum + stCol).Address
  desSht.Range(addr1).Value = Range(addr1).Value
  flag = False
  offR = 0
  stPos = src.Row
  addr1 = Cells(stRow, stCol).Address & ":" & Cells(rowNum, stCol + 1).Address
  desSht.Range(addr1).NumberFormat = "m/d hh:mm:ss"
  While (Count > 0 And src.Row <= Count2 + 1)
   diff = DateDiff("S", ref, src)
   If Abs(diff) <= halfTimeError Then
     flag = True
     Set ref = ref.Offset(1, 0)
     Set src = src.Offset(1, 0)
     Count = Count - 1
   Else
      If flag Then
        endPos = src.Row - 1
        addr1 = Cells(stPos + offR, stCol + 1).Address & ":" & Cells(endPos + offR, colNum).Address
        addr2 = Cells(stPos, stCol + 1).Address & ":" & Cells(endPos, colNum).Address
        desSht.Range(addr1).Value = Range(addr2).Value
        flag = False
        stPos = src.Row
      End If
      If diff > halfTimeError Then
        Set ref = ref.Offset(1, 0)
        offR = offR + 1
        Count = Count - 1
      Else
        Set src = src.Offset(1, 0)
        stPos = src.Row
        offR = offR - 1
      End If
   End If
  Wend
  If flag Then
       endPos = src.Row - 1
       addr1 = Cells(stPos + offR, stCol + 1).Address & ":" & Cells(endPos + offR, colNum).Address
       addr2 = Cells(stPos, stCol + 1).Address & ":" & Cells(endPos, colNum).Address
       desSht.Range(addr1).Value = Range(addr2).Value
  End If
  addr1 = Cells(stRow, stCol).Address & ":" & Cells(rowNum, stCol + 1).Address
  desSht.Range(addr1).NumberFormat = "m/d hh:mm:ss"
  Exit Sub
exitFlag:
 MsgBox Err.Description
  Exit Sub
End Sub


Public Sub setTimeAlign2() '相等
  Dim ref As Range, src As Range
  Dim desSheetName As String
  Dim Count As Long
  Dim stPos As Long, endPos As Long
  stRow = 2
  stCol = 1
  halfTimeError = 0
  Set curSht = ActiveSheet
  desSheetName = "时间对齐数据表"
  On Error Resume Next
  Set desSht = Sheets(desSheetName)
  If desSht Is Nothing Then
    Set desSht = Sheets.Add
    desSht.Name = desSheetName
  Else
    desSht.Range("A1:XFD10000000").Value = ""
  End If
  On Error GoTo exitFlag
  rowNum = Application.WorksheetFunction.CountA(Range(Chr(64 + stCol) & ":" & Chr(64 + stCol)))
  colNum = Application.WorksheetFunction.CountA(Range(stRow & ":" & stRow))
  Count = rowNum - stRow + 1
  Count2 = Application.WorksheetFunction.CountA(Range(Chr(64 + stCol + 1) & ":" & Chr(64 + stCol + 1))) - stRow + 1
  Set ref = Cells(stRow, stCol)
  Set src = Cells(stRow, stCol + 1)
  addr1 = Cells(1, stCol).Address & ":" & Cells(rowNum + stRow, stCol).Address
  desSht.Range(addr1).Value = Range(addr1).Value
  addr1 = Cells(1, stCol).Address & ":" & Cells(stRow - 1, colNum + stCol).Address
  desSht.Range(addr1).Value = Range(addr1).Value
  flag = False
  offR = 0
  stPos = src.Row
  addr1 = Cells(stRow, stCol).Address & ":" & Cells(rowNum, stCol + 1).Address
  desSht.Range(addr1).NumberFormat = "m/d hh:mm:ss"
  While (Count > 0 And src.Row <= Count2 + 1)
   diff = DateDiff("S", ref, src)
   If Abs(diff) <= halfTimeError Then
     flag = True
     Set ref = ref.Offset(1, 0)
     Set src = src.Offset(1, 0)
     Count = Count - 1
   Else
      If flag Then
        endPos = src.Row - 1
        addr1 = Cells(stPos + offR, stCol + 1).Address & ":" & Cells(endPos + offR, colNum).Address
        addr2 = Cells(stPos, stCol + 1).Address & ":" & Cells(endPos, colNum).Address
        desSht.Range(addr1).Value = Range(addr2).Value
        flag = False
        stPos = src.Row
      End If
      If diff > halfTimeError Then
        Set ref = ref.Offset(1, 0)
        offR = offR + 1
        Count = Count - 1
      Else
        Set src = src.Offset(1, 0)
        stPos = src.Row
        offR = offR - 1
      End If
   End If
  Wend
  If flag Then
       endPos = src.Row - 1
       addr1 = Cells(stPos + offR, stCol + 1).Address & ":" & Cells(endPos + offR, colNum).Address
       addr2 = Cells(stPos, stCol + 1).Address & ":" & Cells(endPos, colNum).Address
       desSht.Range(addr1).Value = Range(addr2).Value
  End If
  addr1 = Cells(stRow, stCol).Address & ":" & Cells(rowNum, stCol + 1).Address
  desSht.Range(addr1).NumberFormat = "m/d hh:mm:ss"
  Exit Sub
exitFlag:
 MsgBox Err.Description
  Exit Sub
End Sub

