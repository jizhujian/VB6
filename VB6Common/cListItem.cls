VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListItem"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub FillListItems(ByVal lst As Object, ByVal col As Collection, Optional ByVal ItemTextField As String = "FName", Optional ByVal ItemDataField As String = "FID", Optional ByVal TextIncludeItemData As Boolean)
  Dim i As Long
  With lst
    For i = 1 To col.Count
      If ItemDataField > "" Then
        .AddItem IIf(TextIncludeItemData, "[" & col(i)(ItemDataField) & "]", "") & col(i)(ItemTextField)
        .ItemData(.ListCount - 1) = col(i)(ItemDataField)
      Else
        .AddItem col(i)(ItemTextField)
      End If
    Next
  End With
End Sub

Public Function SelectListItemByItemData(ByVal lst As Object, ByVal Value As Long) As Boolean
  Dim i As Long
  With lst
    For i = 0 To .ListCount - 1
      If .ItemData(i) = Value Then
        SelectListItemByItemData = True
        .ListIndex = i
        Exit For
      End If
    Next
  End With
End Function

Public Function SelectListItemByText(ByVal lst As Object, ByVal Value As String) As Boolean
  Dim i As Long
  With lst
    For i = 0 To .ListCount - 1
      If .List(i) = Value Then
        SelectListItemByText = True
        .ListIndex = i
        Exit For
      End If
    Next
  End With
End Function

