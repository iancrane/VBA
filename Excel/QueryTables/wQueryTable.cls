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

Private omQT As Excel.QueryTable

Public Sub Initialize(ByVal ovQueryTable As Excel.QueryTable)
  Set omQT = ovQueryTable
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
  Name = omQT.Name
End Property

Public Property Get SheetName() As String
  SheetName = omQT.Parent.Name
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
