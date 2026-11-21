VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'This class implements a strongly-typed collection. There are hidden VB procedure attributes,
'so the easiest way to create a new plural collection is to import this class and rename it
'and all instances of "TypedCollectionMember" to the class name you want to use in this collection.

'This class can be found in the VBA repository under Templates

Private mItems As VBA.Collection

Private Sub Class_Initialize()
  Set mItems = New VBA.Collection
End Sub

Private Sub Class_Terminate()
  Set mItems = Nothing
End Sub

Property Get Item(Index As Variant) As TypedCollectionMember
Attribute Item.VB_UserMemId = 0
   Set Item = mItems.Item(Index)
End Property

Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  Set NewEnum = mItems.[_NewEnum]
End Property

Public Sub Add(Item As TypedCollectionMember)
  Call mItems.Add(Item, Item.Key)
End Sub

Public Function Count() As Long
   Count = mItems.Count
End Function

Public Sub Remove(Index As Variant)
   Call mItems.Remove(Index)
End Sub
