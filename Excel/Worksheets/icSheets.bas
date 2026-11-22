Attribute VB_Name = "icSheets"
Option Explicit

Public Function OutputArrayToNewSheet(dataArray As Variant, Optional ByVal destination As String = "A1", Optional ByVal insertAfter As Excel.Worksheet = Nothing) As Excel.Worksheet
  
  Dim ws As Excel.Worksheet
  If insertAfter Is Nothing Then
    Set ws = ThisWorkbook.Sheets.Add()
  Else
    Set ws = ThisWorkbook.Sheets.Add(After:=insertAfter)
  End If

  Dim rows As Integer
  rows = 1 + UBound(dataArray, 1)
  
  Dim cols As Integer
  cols = 1 + UBound(dataArray, 2)
  
  Dim destRange As Excel.range
  Set destRange = ws.range(destination)
  Set destRange = destRange.Resize(rows, cols)
  
  destRange.Value = dataArray

  Set OutputArrayToNewSheet = ws

End Function
