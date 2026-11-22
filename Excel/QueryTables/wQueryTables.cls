VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wQueryTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mItems As VBA.Collection

Public Sub Initialize(Optional ByVal workbook As Excel.workbook = Nothing, Optional ByVal sheet As Excel.Worksheet = Nothing)

  Set mItems = New VBA.Collection
  
  If sheet Is Nothing Then
    Call LoadWorkbookQueries(workbook)
  Else
    Call LoadWorkSheetQueries(sheet)
  End If

End Sub

Private Sub LoadWorkbookQueries(workbook As Excel.workbook)

  Dim olSheet As Excel.Worksheet
  For Each olSheet In workbook.Worksheets
    Call LoadWorkSheetQueries(olSheet)
  Next olSheet

End Sub

Private Sub LoadWorkSheetQueries(sheet As Excel.Worksheet)

  Dim olQuery As Excel.QueryTable
  Dim wQT     As wQueryTable

  For Each olQuery In sheet.QueryTables
    Set wQT = New wQueryTable
    Call wQT.Initialize(olQuery)
    Call mItems.Add(wQT, wQT.Key)
  Next olQuery

End Sub

Property Get Item(Index As Variant) As wQueryTable
  Set Item = mItems.Item(Index)
End Property

Property Get NewEnum() As IUnknown
  Set NewEnum = mItems.[_NewEnum]
End Property

Public Sub Add(Item As wQueryTable)
  Call mItems.Add(Item, Item.Key)
End Sub

Public Function Count() As Long
  Count = mItems.Count
End Function

Public Sub Remove(Index As Variant)
  Call mItems.Remove(Index)
End Sub

Private Sub Class_Initialize()
  Set mItems = New VBA.Collection
End Sub

Private Sub Class_Terminate()
  Set mItems = Nothing
End Sub
