VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetMenuA Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenuA Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemCountA Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemIDA Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_DISABLED = &H2&
Private Const MF_CHECKED = &H8&
Private Const MF_SEPARATOR = &H800&
Private Const WM_COMMAND = &H111

Public Type MenuItemInfoType
  hMenu As Long
  Position As Long
  ItemID As Long
  Caption As String
  Enabled As Boolean
  Checked As Boolean
  IsSeparator As Boolean
End Type

Public Function GetMenu(ByVal hwnd As Long) As Long
  GetMenu = GetMenuA(hwnd)
End Function

Public Function GetSubMenu(ByVal hMenu As Long, ByVal lngPosition As Long) As Long
  GetSubMenu = GetSubMenuA(hMenu, lngPosition)
End Function

Public Function GetMenuItemCount(ByVal hMenu As Long) As Long
  GetMenuItemCount = GetMenuItemCountA(hMenu)
End Function

Public Function GetMenuItemID(ByVal hMenu As Long, ByVal lngPosition As Long) As Long
  GetMenuItemID = GetMenuItemIDA(hMenu, lngPosition)
End Function

Public Function GetMenuItemCaption(ByVal hMenu As Long, ByVal lngPosition As Long) As String
  Dim strCaption As String
  strCaption = String(255, vbNullChar)
  GetMenuString hMenu, lngPosition, strCaption, 255, MF_BYPOSITION
  strCaption = Left(strCaption, InStr(strCaption, vbNullChar) - 1)
  If InStr(strCaption, vbTab) > 0 Then
    strCaption = Left(strCaption, InStr(strCaption, vbTab) - 1)
  End If
  GetMenuItemCaption = strCaption
End Function

Public Function MenuItemIsSeparator(ByVal hMenu As Long, ByVal lngPosition As Long) As Boolean
  Dim lngMenuState As Long
  Dim lngSeparatorState As Long
  lngMenuState = GetMenuState(hMenu, lngPosition, MF_BYPOSITION)
  lngSeparatorState = MF_SEPARATOR Or MF_GRAYED Or MF_DISABLED
  MenuItemIsSeparator = (lngMenuState And lngSeparatorState) = lngSeparatorState
End Function

Public Function GetMenuItemEnabled(ByVal hMenu As Long, ByVal lngPosition As Long) As Boolean
  Dim lngMenuState As Long
  Dim lngDisabledState As Long
  lngMenuState = GetMenuState(hMenu, lngPosition, MF_BYPOSITION)
  lngDisabledState = MF_GRAYED Or MF_DISABLED
  GetMenuItemEnabled = (lngMenuState And lngDisabledState) = 0
End Function

Public Function GetMenuItemChecked(ByVal hMenu As Long, ByVal lngPosition As Long) As Boolean
  Dim lngMenuState As Long
  Dim lngCheckedState As Long
  lngMenuState = GetMenuState(hMenu, lngPosition, MF_BYPOSITION)
  lngCheckedState = MF_CHECKED
  GetMenuItemChecked = (lngMenuState And lngCheckedState) = lngCheckedState
End Function

Public Sub ClickMenuItem(ByVal hwnd As Long, ByVal hMenu As Long)
  SendMessage hwnd, WM_COMMAND, hMenu, ByVal 0&
End Sub

Public Function GetMenuItemInfoByCaption(ByVal hwnd As Long, ParamArray arrMenuItemCaption()) As MenuItemInfoType

  Dim oMenuItemInfo As MenuItemInfoType
  Dim hMenu As Long
  Dim lngCount As Long
  Dim lngPosition As Long
  Dim strCaption As String
  Dim i As Integer
  Dim blnFound As Boolean

  For i = 0 To UBound(arrMenuItemCaption)
    If (i = 0) Then
      hMenu = GetMenu(hwnd)
    Else
      hMenu = GetSubMenu(hMenu, lngPosition)
    End If
    lngCount = GetMenuItemCount(hMenu)
    blnFound = False
    For lngPosition = 0 To lngCount - 1
      strCaption = GetMenuItemCaption(hMenu, lngPosition)
      If UCase(strCaption) = UCase(arrMenuItemCaption(i)) Then
        blnFound = True
        Exit For
      End If
    Next
    If Not blnFound Then
      Exit Function
    End If
  Next
  With oMenuItemInfo
    .hMenu = hMenu
    .Position = lngPosition
    .Caption = strCaption
    .ItemID = GetMenuItemID(hMenu, lngPosition)
    .IsSeparator = MenuItemIsSeparator(hMenu, lngPosition)
    If Not .IsSeparator Then
      .Enabled = GetMenuItemEnabled(hMenu, lngPosition)
      .Checked = GetMenuItemChecked(hMenu, lngPosition)
    End If
  End With
  GetMenuItemInfoByCaption = oMenuItemInfo

End Function
