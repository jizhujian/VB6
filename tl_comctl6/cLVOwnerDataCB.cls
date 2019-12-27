VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVOwnerDataCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IOwnerDataCallback
Private hGIPOld As Long
Private hSIPOld As Long
Private hOCHOld As Long
Private Sub IOwnerDataCallback_GetItemPosition(ByVal iItem As Long, pt As lvundoc.POINT)

End Sub
Private Sub IOwnerDataCallback_SetItemPosition(ByVal iItem As Long, pt As lvundoc.POINT)

End Sub
Private Sub IOwnerDataCallback_GetItemInGroup(ByVal groupindex As Long, ByVal groupWideItemIndex As Long, pTotalItemIndex As Long)
'Debug.Print "GetItemInGroup"
pTotalItemIndex = VLGroups(groupindex).items(groupWideItemIndex)
End Sub
Private Sub IOwnerDataCallback_GetItemGroup(ByVal iItem As Long, ByVal occurenceIndex As Long, pGroupIndex As Long)
'Debug.Print "GetItemGroup"
pGroupIndex = glbRes(iItem).lGroup
End Sub
Private Sub IOwnerDataCallback_GetItemGroupCount(ByVal iItem As Long, pOccurenceCount As Long)
'Debug.Print "GetItemGroupCount"
pOccurenceCount = 1
End Sub
Private Sub IOwnerDataCallback_OnCacheHint(ByVal lviiFirstItem As Long, ByVal lviiFirstGroup As Long, ByVal lviiLastItem As Long, ByVal lviiLastGroup As Long)

End Sub
Private Sub Class_Initialize()

'Dim pVTable As IOwnerDataCallback
'Set pVTable = Me
'hGIPOld = SwapVtableEntry(ObjPtr(pVTable), 4, AddressOf GetItemPositionVB)
'hSIPOld = SwapVtableEntry(ObjPtr(pVTable), 5, AddressOf SetItemPositionVB)
'hOCHOld = SwapVtableEntry(ObjPtr(pVTable), 9, AddressOf OnCacheHintVB)

End Sub



Private Sub Class_Terminate()
''Dim pVTable As IOwnerDataCallback
''Set pVTable = Me
''hGIPOld = SwapVtableEntry(ObjPtr(pVTable), 4, hGIPOld)
''hSIPOld = SwapVtableEntry(ObjPtr(pVTable), 5, hSIPOld)
''hOCHOld = SwapVtableEntry(ObjPtr(pVTable), 9, hOCHOld)

End Sub

