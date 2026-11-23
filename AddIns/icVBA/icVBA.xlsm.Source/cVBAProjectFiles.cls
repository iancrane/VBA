VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "cVBAProjectFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private smModule As String

Private SharedOnly As Boolean
Private Project As cVBAProject

Public Function Export(ByVal ovWorkbook As Excel.Workbook) As epdReturnCode

	Call Initialize(ovWorkbook, False)

	Export = Validate()
	If Export <> eReturnCodeSuccess Then
		Exit Function
	End If

	Call FileSystem.CreatePath(Project.ProjectDirectory)
	Export = ProcessProjectComponents(False)

End Function

Public Function Import(ByVal ovWorkbook As Excel.Workbook, ByVal bvSharedOnly As Boolean) As epdReturnCode

	Call Initialize(ovWorkbook, bvSharedOnly)

	Import = Validate()
	If Import <> eReturnCodeSuccess Then
		Exit Function
	End If

	Import = ProcessProjectComponents(True)

End Function

Private Sub Initialize(ByVal ovWorkbook As Excel.Workbook, ByVal bvSharedOnly As Boolean)

	Set Project = NewVBAProject(ovWorkbook)
	SharedOnly = bvSharedOnly

End Sub

Private Function NewVBAProject(ByVal ovWorkbook As Excel.Workbook) As cVBAProject

	Set NewVBAProject = New cVBAProject
	Call NewVBAProject.Initialize(ovWorkbook)

End Function

Private Function ProcessProjectComponents(ByVal bvIsImport As Boolean) As epdReturnCode

	On Error GoTo HandleErr

	Dim olComponents	As VBComponents
	Dim olComponent		As VBComponent
	Dim componentList	As New Collection

	Set olComponents = Project.Workbook.VBProject.VBComponents
	For Each olComponent In olComponents
		Call componentList.Add(olComponent)
	Next olComponent

	Dim result As Long
	For Each olComponent In componentList
		result = ProcessProjectComponent(bvIsImport, olComponent, olComponents)
		If result <> eReturnCodeSuccess Then ProcessProjectComponents = eReturnCodeFailed
	Next olComponent
	Exit Function

HandleErr:

	ProcessProjectComponents = Utility.HandleUnexpectedError(smModule, "ProcessProjectComponents", Utility.ErrInfoList("bvIsImport", bvIsImport))
	Exit Function

End Function

Private Function ProcessProjectComponent(ByVal bvIsImport As Boolean, ByVal ovComponent As VBComponent, ByVal ovComponents As VBComponents) As Long

	On Error GoTo HandleErr

	Dim slPath As String
	If ShouldProcessComponent(ovComponent, slPath) = False Then
		Exit Function
	End If

	If bvIsImport Then
		If CanImportComponent(ovComponent, slPath) Then
			ProcessProjectComponent = ImportProjectComponent(ovComponent, ovComponents, slPath)
		End If
	Else
		ProcessProjectComponent = ExportProjectComponent(ovComponent, slPath)
	End If

	Exit Function

HandleErr:

	ProcessProjectComponent = Utility.HandleUnexpectedError(smModule, "ProcessProjectComponent", Utility.ErrInfoList("bvIsImport", bvIsImport, "slPath", slPath))
	Exit Function

End Function

Private Function CanImportComponent(ByVal component As VBComponent, ByVal filePath As String) As Boolean

	If Not Utility.FileExists(filePath) Then Exit Function

	'Keep the modules that run the importer loaded while it updates this workbook.
	If Project.Workbook Is ThisWorkbook Then
		Select Case component.Name
			Case "basExportVBAProject", "cVBAProject", "cVBAProjectFiles", "cFileSystem", "cUtility"
				Exit Function
		End Select
	End If

	CanImportComponent = True

End Function

Private Function ImportProjectComponent(ovComponent As VBComponent, ovComponents As VBComponents, ByVal svFilePath As String) As Long

	On Error GoTo HandleErr

	Dim componentName As String
	componentName = ovComponent.Name

	Dim importedComponent As VBComponent
	Set importedComponent = ovComponents.Import(svFilePath)
	Call ovComponents.Remove(ovComponent)
	importedComponent.Name = componentName
	Exit Function

HandleErr:

	ImportProjectComponent = Utility.HandleUnexpectedError(smModule, "ImportProjectComponent", svFilePath)
	If Not importedComponent Is Nothing Then
		On Error Resume Next
		If Not ovComponents(componentName) Is Nothing Then Call ovComponents.Remove(importedComponent)
	End If

End Function

Private Function ExportProjectComponent(ovComponent As VBComponent, ByVal svFilePath As String) As Long

	On Error GoTo HandleErr

	'If the file exists, overwrite the existing file
	If Utility.FileExists(svFilePath) Then
		Kill svFilePath
	End If

	'Export the ovComponent
	Call ovComponent.Export(svFilePath)
	Exit Function

HandleErr:

	ExportProjectComponent = Utility.HandleUnexpectedError(smModule, "ExportProjectComponent", svFilePath)
	Exit Function

End Function

Private Function ShouldProcessComponent(ByVal ovComponent As VBComponent, ByRef srPath As String) As Boolean

	ShouldProcessComponent = ShouldProcessComponentEx(ovComponent, srPath)
	Call TraceProcessComponent(ShouldProcessComponent, srPath)
End Function

Private Function ShouldProcessComponentEx(ByVal ovComponent As VBComponent, ByRef srPath As String) As Boolean

	srPath = ComponentFileName(ovComponent)

	If SharedOnly Then
		ShouldProcessComponentEx = IsShared(srPath)
	Else
		ShouldProcessComponentEx = ovComponent.Type <> vbext_ct_Document  'Can't process items listed under "Microsoft Excel Objects", like ThisWorkbook or Sheet1, etc
	End If

	'Return this either way for (logging)
	srPath = GetComponentSaveAsPath(srPath)

End Function

Private Function TraceProcessComponent(ByVal vvResult As Variant, ByVal svPath As String)

	Call Utility.Trace(smModule, "ShouldProcessComponent", vvResult & vbTab & svPath)
End Function

Private Function GetComponentSaveAsPath(ByVal svFileName As String) As String

	GetComponentSaveAsPath = Utility.QualifyFilePath(svFileName, ComponentDirectory(svFileName))

End Function

Private Function ComponentDirectory(svFileName As String) As String

	If IsShared(svFileName) Then
		ComponentDirectory = Project.SharedDirectory
	Else
		ComponentDirectory = Project.ProjectDirectory
	End If

End Function

Private Function ComponentFileName(ovComponent As VBComponent) As String

	ComponentFileName = ovComponent.Name & Ext(ovComponent)
End Function

Private Function IsShared(ByVal svFileName As String) As Boolean

	Select Case svFileName
		Case "basDeveloper.bas", "basDrillDown.bas", "basFunkyXL.bas", "basXL.bas", "cDDParameter.cls", "cDDWorksheetWrapper.cls", "cEvent.cls", "cEvents.cls", "cExcelLinks.cls", "cIEventsOwner.cls", "cIExcelLinksOwner.cls", "cUtility.cls", "cXlaSettings.cls", "icEventArgs.cls", "wVBAComponent.cls"
			IsShared = True
	End Select

End Function

Private Function Ext(ovComponent As VBComponent) As String

	Select Case ovComponent.Type
		Case vbext_ct_StdModule:  Ext = ".bas"
		Case vbext_ct_MSForm:     Ext = ".frm"
		Case Else:                Ext = ".cls"
	End Select

End Function

Private Function Validate() As epdReturnCode

	On Error GoTo HandleErr

	'Prompt user that the VBProject is locked from viewing
	If Project.Workbook.VBProject.Protection = vbext_pp_locked Then
		Call MsgBox("The workbook " & Project.Workbook.Name & " VBA Project is locked.", vbInformation)
		Validate = eReturnCodeFailed
		Exit Function
	End If

	Exit Function

HandleErr:

	Validate = Utility.HandleUnexpectedError(smModule, "Validate")
	Exit Function

End Function

Private Sub Class_Initialize()

	smModule = TypeName(Me)
End Sub
