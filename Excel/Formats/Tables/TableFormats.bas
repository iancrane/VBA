Attribute VB_Name = "TableFormats"
Option Explicit


Private Const DefaultBandingMarker As String = "QTMeta.AutoFormat"

Public Sub Autoformat(Optional ByVal queryTable As Excel.QueryTable = Nothing, _
											Optional ByVal resultRange As Excel.Range = Nothing, _
											Optional ByVal bandingMarker As String = DefaultBandingMarker)

	Dim table As Excel.ListObject
	If Not queryTable Is Nothing Then Set table = QueryListObject(queryTable)

	Dim result	As Excel.Range
	Dim body		As Excel.Range
	Dim header	As Excel.Range

	If Not table Is Nothing Then
		Set body = table.DataBodyRange
		If table.ShowHeaders Then Set header = table.HeaderRowRange
	ElseIf Not queryTable Is Nothing Then
		Set result = queryTable.ResultRange
		If result Is Nothing Then Exit Sub

		Set body = result
		If queryTable.FieldNames Then
			Dim destinationRow As Long
			destinationRow = queryTable.Destination.Row
			Set header = result.Worksheet.Cells(destinationRow, result.Column).Resize(1, result.Columns.Count)

			'Some QueryTables include the field names in ResultRange.
			If result.Row = destinationRow Then Set body = RowsBelowHeader(result)
		End If
	ElseIf Not resultRange Is Nothing Then
		Set header = resultRange.Rows(1)
		Set body = RowsBelowHeader(resultRange)
	End If

	Set result = body
	If Not header Is Nothing Then
		Set result = header
		If Not body Is Nothing Then Set result = header.Resize(body.Rows.Count + 1)
	End If
	If result Is Nothing Then Exit Sub

	Dim app As Excel.Application
	Set app = result.Application

	Dim wasUpdating As Boolean
	wasUpdating = app.ScreenUpdating
	On Error GoTo RestoreScreenUpdating
	app.ScreenUpdating = False

	Call AutoSizeColumns(result)
	Call ApplyBanding(body, table, bandingMarker)
	Call ApplyBorders(result, header)
	If Not header Is Nothing Then Call FormatHeader(header)

	app.ScreenUpdating = wasUpdating
	Exit Sub

RestoreScreenUpdating:
	app.ScreenUpdating = wasUpdating
	Call Err.Raise(Err.Number, Err.Source, Err.Description)
End Sub

Public Sub AutoSizeColumns(ByVal range As Excel.Range)

	If range Is Nothing Then Exit Sub

	If range.Areas.Count = 1 Then
		Call range.Columns.AutoFit()
		Exit Sub
	End If

	'Keep the widest fit when separate areas share a worksheet column.
	Dim widths() As Double
	ReDim widths(1 To range.Worksheet.Columns.Count)

	Dim area				As Excel.Range
	Dim column			As Excel.Range
	Dim index				As Long
	Dim fittedWidth	As Double
	For Each area In range.Areas
		Call area.Columns.AutoFit()
		For Each column In area.Columns
			index = column.Column
			fittedWidth = column.ColumnWidth
			If fittedWidth < widths(index) Then
				column.ColumnWidth = widths(index)
			Else
				widths(index) = fittedWidth
			End If
		Next column
	Next area
End Sub

Public Sub ApplyBanding(ByVal body As Excel.Range, _
	Optional ByVal table As Excel.ListObject = Nothing, _
	Optional ByVal bandingMarker As String = DefaultBandingMarker)

	If body Is Nothing Then Exit Sub

	If table Is Nothing Then
		Call ApplyRowBanding(body, bandingMarker)
	Else
		Call ApplyNativeRowBanding(table, body, bandingMarker)
	End If
End Sub

Public Sub ApplyBorders(ByVal result As Excel.Range, Optional ByVal header As Excel.Range = Nothing)

	'When supplied, header is the first row of result.
	If result Is Nothing Then Exit Sub

	Dim body As Excel.Range
	Set body = result
	If Not header Is Nothing Then Set body = RowsBelowHeader(result)

	Dim columnCount As Long
	columnCount = result.Columns.Count

	Dim gridGray				As Long
	Dim upperEdgeGray		As Long
	Dim headerGridGray	As Long
	Dim lowerEdgeGray		As Long
	gridGray = RGB(224, 224, 224)
	upperEdgeGray = RGB(173, 173, 173)
	headerGridGray = RGB(128, 128, 128)
	lowerEdgeGray = RGB(79, 79, 79)

	If Not body Is Nothing Then
		If body.Rows.Count > 1 Then Call SetTableBorder(body, xlInsideHorizontal, gridGray)
		If columnCount > 1 Then Call SetTableBorder(body, xlInsideVertical, gridGray)
	End If

	Call SetTableBorder(result, xlEdgeTop, upperEdgeGray)
	Call SetTableBorder(result, xlEdgeLeft, upperEdgeGray)
	Call SetTableBorder(result, xlEdgeBottom, lowerEdgeGray)
	Call SetTableBorder(result, xlEdgeRight, lowerEdgeGray)

	If Not header Is Nothing Then
		If columnCount > 1 Then Call SetTableBorder(header, xlInsideVertical, headerGridGray)
		Call SetTableBorder(header, xlEdgeBottom, lowerEdgeGray)
	End If
End Sub

Private Function RowsBelowHeader(ByVal result As Excel.Range) As Excel.Range

	Dim rowCount As Long
	rowCount = result.Rows.Count
	If rowCount > 1 Then Set RowsBelowHeader = result.Resize(rowCount - 1).Offset(1)
End Function

Private Sub FormatHeader(ByVal header As Excel.Range)

	header.Interior.Color = RGB(156, 194, 229)
	With header.Font
		.Bold = True
		.Color = RGB(31, 55, 82)
	End With

	Dim sheet As Excel.Worksheet
	Set sheet = header.Worksheet

	Dim book As Excel.Workbook
	Set book = sheet.Parent

	Dim sheetWindow As Excel.Window
	For Each sheetWindow In book.Windows
		If sheetWindow.Visible Then
			If sheetWindow.ActiveSheet Is sheet Then Call FreezeTableHeader(sheetWindow, header.Row)
		End If
	Next sheetWindow
End Sub

Private Sub FreezeTableHeader(ByVal sheetWindow As Excel.Window, ByVal headerRow As Long)
  If sheetWindow Is Nothing Then Exit Sub

  With sheetWindow
    If .FreezePanes And .SplitColumn = 0 Then
      If .Panes(1).ScrollRow + .SplitRow - 1 = headerRow Then Exit Sub
    End If

    .FreezePanes = False
    .Split = False
    .ScrollRow = 1

    'Keep room to scroll when the header is near or below the bottom of the window.
    If headerRow >= .VisibleRange.Rows.Count - 1 Then
      .ScrollRow = headerRow
      .SplitRow = 1
    Else
      .SplitRow = headerRow
    End If
    .FreezePanes = True
  End With
End Sub

Private Function QueryListObject(ByVal queryTable As Excel.QueryTable) As Excel.ListObject
  'Excel raises error 1004 for standalone queries.
  On Error Resume Next
  Set QueryListObject = queryTable.ListObject
  On Error GoTo 0
End Function

Private Sub ApplyNativeRowBanding(ByVal table As Excel.ListObject, ByVal body As Excel.Range, ByVal bandingMarker As String)

  Const StylePrefix As String = "icVisions.QTMeta.Banded."

  Dim currentStyle As Excel.TableStyle
  Dim currentStyleName As String
  Dim styleName As String
  styleName = StylePrefix & "(none)"

  'TableStyle returns a string when the table has no style.
  If IsObject(table.TableStyle) Then
    Set currentStyle = table.TableStyle
    currentStyleName = currentStyle.Name
    styleName = StylePrefix & currentStyleName
  End If

  If Left$(currentStyleName, Len(StylePrefix)) <> StylePrefix Then
    Dim book As Excel.Workbook
    Set book = body.Worksheet.Parent

    Dim bandedStyle As Excel.TableStyle
    On Error Resume Next
    Set bandedStyle = book.TableStyles(styleName)
    On Error GoTo 0

    If bandedStyle Is Nothing Then
      If currentStyle Is Nothing Then
        Set bandedStyle = book.TableStyles.Add(styleName)
      Else
        Set bandedStyle = currentStyle.Duplicate(styleName)
      End If
      With bandedStyle.TableStyleElements(xlRowStripe1)
        .StripeSize = 1
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(221, 235, 247)
      End With
      With bandedStyle.TableStyleElements(xlRowStripe2)
        .StripeSize = 1
        .Interior.Pattern = xlSolid
        .Interior.Color = vbWhite
      End With
    End If

    table.TableStyle = bandedStyle.Name
  End If
  If Not table.ShowTableStyleRowStripes Then table.ShowTableStyleRowStripes = True

  'Direct fills hide the native stripes, including the old formatter's white base.
  Dim fillPattern As Variant
  fillPattern = body.Interior.Pattern
  If IsNull(fillPattern) Or fillPattern <> xlPatternNone Then body.Interior.Pattern = xlPatternNone

  Dim banding As Excel.FormatCondition
  Set banding = FindRowBanding(body, bandingMarker)
  If Not banding Is Nothing Then Call banding.Delete()
End Sub

Private Function FindRowBanding(ByVal body As Excel.Range, ByVal bandingMarker As String) As Excel.FormatCondition

  Dim condition As Object
  For Each condition In body.FormatConditions
    If condition.Type = xlExpression Then
      If InStr(1, condition.Formula1, bandingMarker, vbBinaryCompare) > 0 Then
        Set FindRowBanding = condition
        Exit Function
      End If
    End If
  Next condition
End Function

Private Sub ApplyRowBanding(ByVal body As Excel.Range, ByVal bandingMarker As String)

  Dim bandingFormula As String
  bandingFormula = "=MOD(ROW()-" & body.Row & ",2)=N(""" & bandingMarker & """)"

  Dim banding As Excel.FormatCondition
  Set banding = FindRowBanding(body, bandingMarker)

  If banding Is Nothing Then
    Set banding = body.FormatConditions.Add(Type:=xlExpression, Formula1:=bandingFormula)
  Else
    If banding.Formula1 <> bandingFormula Then Call banding.Modify(xlExpression, Formula1:=bandingFormula)
    If banding.AppliesTo.Address <> body.Address Then Call banding.ModifyAppliesToRange(body)
  End If

  body.Interior.Color = vbWhite
  With banding
    If .Priority <> 1 Then Call .SetFirstPriority()
    .StopIfTrue = False
    .Interior.Color = RGB(221, 235, 247)
  End With
End Sub

Private Sub SetTableBorder(ByVal target As Excel.Range, ByVal edge As Excel.XlBordersIndex, ByVal color As Long)
  With target.Borders(edge)
    .LineStyle = xlContinuous
    .Weight = xlThin
    .Color = color
  End With
End Sub
