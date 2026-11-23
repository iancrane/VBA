VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wQueryTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private omQT		As Excel.QueryTable
Private omTable	As Excel.ListObject

Public Sub Initialize(ByVal ovQueryTable As Excel.QueryTable)

	Set omQT = ovQueryTable
	Set omTable = Nothing

	'Standalone queries raise an error when asked for a backing table.
	On Error Resume Next
	Set omTable = omQT.ListObject
	On Error GoTo 0
End Sub

Public Property Get Key() As String
  Key = Me.SheetName & "!" & Me.Name
End Property

Public Property Get ResultRange() As String
  ResultRange = omQT.ResultRange.Address(False, False)
End Property

Public Property Get Connection() As String
  Connection = omQT.Connection
End Property

Public Property Get Name() As String

	If omTable Is Nothing Then
		Name = omQT.Name
	Else
		'Some table-backed queries expose their name only on the ListObject.
		On Error Resume Next
		Name = omQT.Name
		On Error GoTo 0
		If Len(Name) = 0 Then Name = omTable.Name
	End If
End Property

Public Property Get SheetName() As String

	If omTable Is Nothing Then
		SheetName = omQT.Parent.Name
	Else
		SheetName = omTable.Parent.Name
	End If
End Property

Public Property Get CommandText() As String
  CommandText = omQT.CommandText
End Property

Public Sub FillMetadataRow(table, rowIndex)

  table(rowIndex, 0) = Me.SheetName
  table(rowIndex, 1) = Me.Name
  table(rowIndex, 2) = Me.ResultRange
  table(rowIndex, 3) = Me.Connection
  table(rowIndex, 4) = Me.CommandText
  
End Sub
