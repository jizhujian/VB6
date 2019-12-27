VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLVFooterCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IListViewFooterCallback

Public Sub IListViewFooterCallback_OnButtonClicked(ByVal itemIndex As Long, ByVal lParam As Long, pRemoveFooter As Long)
Debug.Print "Get footer button click, index=" & itemIndex & ",lparam=" & lParam
HandleLVFooterClick itemIndex, lParam
End Sub
Public Sub IListViewFooterCallback_OnDestroyButton(ByVal itemIndex As Long, ByVal lParam As Long)

End Sub
