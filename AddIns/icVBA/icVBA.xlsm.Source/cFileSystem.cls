VERSION 1.0 CLASS
BEGIN
	MultiUse = -1  'True
END
Attribute VB_Name = "cFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private smModule As String

Public Function FolderExists(ByVal svFolder As String) As Boolean

	'On some machines, passing a blank is not caught -- do an explicit check here
	If svFolder = "" Then Exit Function

	'If the path is UNC and the root of a share, make sure it ends with a backslash so it will work with Dir
	If svFolder Like "\\*\*" And Not (svFolder Like "\\*\*\*") Then
		svFolder = svFolder & "\"
	End If

	'If no directory entry found with this name, get out
	If Dir(svFolder, vbReadOnly Or vbHidden Or vbSystem Or vbDirectory) = "" Then
		Exit Function
	End If

   'Check to see if directory entry is a directory.
	FolderExists = (GetAttr(svFolder) And vbDirectory) <> 0

End Function

Public Sub CreatePath(ByVal svPath As String)

	If FolderExists(svPath) Then Exit Sub

	Dim ilCurrentPosition As Integer
	ilCurrentPosition = SeekRoot(svPath)

		'Make sure the path ends with a backslash
	If Right$(svPath, 1) <> "\" Then
		svPath = svPath & "\"
	End If

	Call Trace("CreatePath", svPath)

	'Work our way down the path one directory at a time
	Dim slPathSoFar As String
	ilCurrentPosition = MoveNext(ilCurrentPosition, svPath)
	Do While ilCurrentPosition > 0

		'Extract the string representing the path up to this directory
		slPathSoFar = Left$(svPath, ilCurrentPosition - 1)

		'If the directory does not yet exist, create it now
		If Not FolderExists(slPathSoFar) Then
			MkDir slPathSoFar
		End If

		'Move to the next directory
		ilCurrentPosition = MoveNext(ilCurrentPosition, svPath)

	Loop

End Sub

Private Function FindNth(ByVal ivOrdinalPosition As Integer, ByVal svPath As String, ByVal svToken As String) As Integer

	Dim ili As Integer, ilCurrent As Integer
	For ili = 1 To ivOrdinalPosition
		ilCurrent = InStr(ilCurrent + 1, svPath, svToken)
	Next ili

	FindNth = ilCurrent

End Function

Private Function MoveNext(ByVal ivCurrentPosition As Integer, ByVal svPath As String) As Integer

	MoveNext = InStr(ivCurrentPosition + 1, svPath, "\")
End Function

Private Function SeekRoot(ByVal svPath As String) As Integer

	Dim ilnSlashes As Integer
	If IsUncFormat(svPath) Then
		ilnSlashes = 3
	Else
		ilnSlashes = 1
	End If

	SeekRoot = FindNth(ilnSlashes, "\", svPath)

End Function

Public Function IsUncFormat(svPath As String)

	'Weak implementation - can fix later :)
	IsUncFormat = svPath Like "\\*\*"

End Function

Private Sub Trace(svProcedure As String, Optional svInfo As String = "")

	Call Utility.Trace(smModule, svProcedure, svInfo)
End Sub

Private Sub Class_Initialize()

	smModule = TypeName(Me)
End Sub
