Attribute VB_Name = "basExportVBAProject"
Option Explicit
Private Const smcModule As String = "basExportVBAProject"
Public Const spcAppTitle As String = "icVBA"

Public Enum epdReturnCode
	eReturnCodeSuccess = 0
	eReturnCodeFailed = -1
	eReturnCodeCancel = 32755
End Enum

Public FileSystem As New cFileSystem
Public Utility As New cUtility

Public Sub ExportLR()

	Call ExportNamedVBAProject("LR.xls")
End Sub

Public Sub ImportLR()

	Call ImportNamedVBAProject("LR.xls")
End Sub

Public Sub ExportOC()

	Call ExportNamedVBAProject("OfficeConnector.xls")
End Sub

Public Sub ImportOC()

	Call ImportNamedVBAProject("OfficeConnector.xls")
End Sub

Public Sub ExportQTFix()

	Call ExportNamedVBAProject("Excel 2016 QueryTable Fix.xlsm")
End Sub

Public Sub ImportQTFix()

	Call ImportNamedVBAProject("Excel 2016 QueryTable Fix.xlsm")
End Sub

Public Sub ExportThisVBAProject()

	Call ExportVBAProject(ThisWorkbook)
End Sub

Public Sub ImportThisVBAProject(Optional ByVal bvSharedOnly As Boolean = True)

	Call ImportVBAProjectComponents(ThisWorkbook, bvSharedOnly)
End Sub

Public Sub ExportCurrentVBAProject()

	Call ExportVBAProject(ActiveWorkbook)
End Sub

Public Sub ImportCurrentVBAProject(Optional ByVal bvSharedOnly As Boolean = True)

	Call ImportVBAProjectComponents(ActiveWorkbook, bvSharedOnly)
End Sub

'Example syntax: Call ExportNamedVBAProject("OfficeConnector.xls")
Public Sub ExportNamedVBAProject(ByVal svWorkbookName As String)

	Call ExportVBAProject(GetNamedWorkbook(svWorkbookName))
End Sub

'Example syntax: Call ImportNamedVBAProject("OfficeConnector.xls")
Public Sub ImportNamedVBAProject(ByVal svWorkbookName As String, Optional ByVal bvSharedOnly As Boolean = True)

	Call ImportVBAProjectComponents(GetNamedWorkbook(svWorkbookName), bvSharedOnly)
End Sub

Public Sub ExportVBAProject(ByVal ovWorkbook As Excel.Workbook)

	If ovWorkbook Is Nothing Then Exit Sub
	Call Trace("ExportVBAProject", ovWorkbook.Name)

	Dim projectFiles As New cVBAProjectFiles
	Call projectFiles.Export(ovWorkbook)

End Sub

Public Sub ImportVBAProjectComponents(ByVal ovWorkbook As Excel.Workbook, ByVal bvSharedOnly As Boolean)

	If ovWorkbook Is Nothing Then Exit Sub
	Call Trace("ImportVBAProjectComponents", ovWorkbook.Name & vbTab & "(" & "SharedOnly=" & bvSharedOnly & ")")

	Dim projectFiles As New cVBAProjectFiles
	Call projectFiles.Import(ovWorkbook, bvSharedOnly)

End Sub

Private Function GetNamedWorkbook(ByVal svName As String) As Excel.Workbook

	On Error Resume Next
	Set GetNamedWorkbook = Application.Workbooks(svName)
	On Error GoTo 0
	If Not GetNamedWorkbook Is Nothing Then Exit Function

	Dim addIn As Excel.AddIn
	For Each addIn In Application.AddIns
		If StrComp(addIn.Name, svName, vbTextCompare) = 0 Or StrComp(addIn.Title, svName, vbTextCompare) = 0 Then
			If addIn.IsOpen Then
				Set GetNamedWorkbook = Application.Workbooks(addIn.Name)
				Exit Function
			End If
		End If
	Next addIn

	Call MsgBox("The workbook " & svName & " is not open.", vbInformation)

End Function

Private Sub Trace(svProcedure As String, Optional svInfo As String = "")

	Call Utility.Trace(smcModule, svProcedure, svInfo)
End Sub
