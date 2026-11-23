VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "cVBAProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ProjectDirectory As String
Public SharedDirectory As String

Private RootDirectory As String

Private Enum emdProjectClassification
	AddIn = 0
	Distribution = 1
	Services = 2
End Enum
Private Classification As emdProjectClassification

Private WorkbookName As String
Private mWorkbook As Excel.Workbook

Public Sub Initialize(ByVal ovWorkbook As Excel.Workbook)

	Set mWorkbook = ovWorkbook
	WorkbookName = Utility.ExtractFileName(Workbook.Name, False)

	Classification = GetClassification()

	RootDirectory = "T:\git\VBA\" & ClassificationFolderName & "\"
	SharedDirectory = SubDir("Shared")
	If Workbook Is ThisWorkbook Or Not FileSystem.FolderExists(SharedDirectory) Then
		SharedDirectory = Utility.QualifyFilePath("Shared", Utility.ExtractPath(ThisWorkbook.Path))
	End If
	ProjectDirectory = GetProjectDirectory()

End Sub

Private Function GetProjectDirectory() As String


	If Classification = AddIn And FileSystem.FolderExists(RootDirectory) Then
		GetProjectDirectory = GetAddInSourceDirectory()
	Else
		GetProjectDirectory = GetLocalSourceDirectory()
	End If

End Function

Private Function GetAddInSourceDirectory() As String

	Select Case WorkbookName
		Case "LR":                    GetAddInSourceDirectory = SubDir("LR")
		Case "OfficeConnector":       GetAddInSourceDirectory = SubDir("OC")
		Case "OfficeConnectorImport": GetAddInSourceDirectory = SubDir("OCI")
		Case "icVisionsLogHelper":     GetAddInSourceDirectory = SubDir("LogHelper")
		Case Else:                    GetAddInSourceDirectory = SubDir(WorkbookName)
	End Select

End Function

Private Function GetLocalSourceDirectory() As String

	Dim fullNameSource As String
	fullNameSource = Utility.QualifyFilePath(Workbook.Name & ".Source", Workbook.Path)

	If Workbook Is ThisWorkbook Or FileSystem.FolderExists(fullNameSource) Then
		GetLocalSourceDirectory = fullNameSource
	Else
		GetLocalSourceDirectory = Utility.QualifyFilePath(WorkbookName & ".Source", Workbook.Path)
	End If

End Function

Private Function SubDir(ByVal svAdditionalFolders As String) As String

	SubDir = Utility.QualifyFilePath(svAdditionalFolders, RootDirectory)
End Function

Private Function GetClassification() As emdProjectClassification

	Select Case WorkbookName
		Case "LR", "OfficeConnector", "OfficeConnectorImport", "icVisionsLogHelper"
			GetClassification = AddIn
		Case Else
			If InStr(1, Workbook.Path & "\", "\Distribution\VB6\", vbTextCompare) > 0 Then
				GetClassification = Distribution
			Else
				GetClassification = Services
			End If
	End Select

End Function

Private Function ClassificationFolderName() As String

	Select Case Classification
		Case emdProjectClassification.AddIn:        ClassificationFolderName = "AddIns"
		Case emdProjectClassification.Distribution: ClassificationFolderName = "Distribution"
		Case Else:                                  ClassificationFolderName = "Services"
	End Select

End Function

Public Property Get Workbook() As Excel.Workbook
	Set Workbook = mWorkbook
End Property
