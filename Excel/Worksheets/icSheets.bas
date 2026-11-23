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


Public Sub DeleteEmptySheets()

	Dim book	As Excel.Workbook
	Dim app		As Excel.Application
	Set book = ThisWorkbook
	Set app = book.Application

	Dim visibleSheetCount	As Long
	Dim sheet							As Object
	For Each sheet In book.Sheets
		If sheet.Visible = xlSheetVisible Then visibleSheetCount = visibleSheetCount + 1
	Next sheet

	Dim worksheets					As Excel.Sheets
	Dim worksheet						As Excel.Worksheet
	Dim index								As Long
	Dim isVisible						As Boolean
	Dim wasDisplayingAlerts	As Boolean
	Set worksheets = book.Worksheets

	wasDisplayingAlerts = app.DisplayAlerts
	On Error GoTo RestoreDisplayAlerts
	app.DisplayAlerts = False

	For index = worksheets.Count To 1 Step -1
		Set worksheet = worksheets(index)
		isVisible = worksheet.Visible = xlSheetVisible
		If Not isVisible Or visibleSheetCount > 1 Then
			If app.WorksheetFunction.CountA(worksheet.UsedRange) = 0 Then
				Call DeleteWorksheet(worksheet)
				If isVisible Then visibleSheetCount = visibleSheetCount - 1
			End If
		End If
	Next index

	app.DisplayAlerts = wasDisplayingAlerts
	Exit Sub

RestoreDisplayAlerts:
	app.DisplayAlerts = wasDisplayingAlerts
	Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

Private Sub DeleteWorksheet(ByVal worksheet As Excel.Worksheet)

	If worksheet.Visible <> xlSheetVeryHidden Then
		Call worksheet.Delete()
		Exit Sub
	End If

	'Excel can delete a hidden sheet, but rejects a very hidden sheet.
	worksheet.Visible = xlSheetHidden
	On Error GoTo RestoreVisibility
	Call worksheet.Delete()
	Exit Sub

RestoreVisibility:
	worksheet.Visible = xlSheetVeryHidden
	Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub
