VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const toolbarName As String = "icVBA"

Public Sub CreateToolbar()

	Dim toolbar	As Office.CommandBar
	Dim button	As Office.CommandBarButton

	Set toolbar = GetToolbar()
	If Not toolbar Is Nothing Then Exit Sub

	Set toolbar = Application.CommandBars.Add(Name:=toolbarName, Position:=msoBarTop, Temporary:=True)
	Set button = toolbar.Controls.Add(Type:=msoControlButton, Temporary:=True)

	With button
		.Caption = "Export VBA Project"
		.TooltipText = "Export VBA Project source files"
		.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!ExportCurrentVBAProject"
		.Style = msoButtonIcon
		.Tag = toolbarName
		.Picture = frmPackrat.imgPasteNewLog.Picture
		.Mask = frmPackrat.imgPasteNewLogMask.Picture
	End With

	toolbar.Visible = True
End Sub

Public Sub DestroyToolbar()

	Dim toolbar As Office.CommandBar
	Set toolbar = GetToolbar()
	If Not toolbar Is Nothing Then Call toolbar.Delete()
End Sub

Private Function GetToolbar() As Office.CommandBar

	On Error Resume Next
	Set GetToolbar = Application.CommandBars(toolbarName)
End Function
