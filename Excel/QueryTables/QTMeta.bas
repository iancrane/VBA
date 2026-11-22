Attribute VB_Name = "QTMeta"
Option Explicit

Public Sub ListQueries()
  
  Dim wQTs As New wQueryTables
  Call wQTs.Initialize(ThisWorkbook)
  
  Dim outputTable() As Variant
  ReDim Preserve outputTable(wQTs.Count, 4)
	
  'Header row
  outputTable(0, 0) = "Parent Sheet"
  outputTable(0, 1) = "QueryTable Name"
  outputTable(0, 2) = "Result Range"
  outputTable(0, 3) = "Connection"
  outputTable(0, 4) = "CommandText" 
  
  Dim i As Integer
  For i = 1 To wQTs.Count
    Call wQTs.Item(i).FillMetadataRow(outputTable, i)
  Next i
  
	Dim ws As Excel.Worksheet
  Set ws = icSheets.OutputArrayToNewSheet(outputTable)
  
End Sub
