Attribute VB_Name = "modApi32"
'****************************************************************************
'
'枕善居汉化收藏整理
'发布日期：05/07/05
'描  述：组件属性窗口控件 Ver1.0
'网  站：http://www.codesky.net/
'
'
'****************************************************************************
Option Explicit

Private Const SM_CXBORDER = 5
Private Const SM_CXVSCROLL = 2
Private Const WS_VSCROLL = &H200000
Private Const GWL_STYLE = (-16)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Declare Function ApiSetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long

Public Function ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oThis As Object

    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oThis, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oThis
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oThis, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but this will be your code rather than the uncounted reference!
End Function

Public Function PtrFromObject(ByRef oThis) As Long
    ' Return the pointer to this object:
    PtrFromObject = ObjPtr(oThis)
End Function

Function TrimNull(Item As String) As String
    Dim pos As Integer

    'double check there is a chr$(0) in the string
    pos = InStr(Item, Chr$(0))
    If pos Then
        TrimNull = Left$(Item, pos - 1)
    Else
        TrimNull = Item
    End If
End Function

Public Sub StopFlicker(ByVal lHwnd As Long)
    Dim lRet As Long
    ' object will not flicker - just be blank
    lRet = LockWindowUpdate(lHwnd)
End Sub

Public Sub Release()
    Dim lRet As Long
    lRet = LockWindowUpdate(0)
End Sub

Function ScrollBarVisible(hWndA As Long) As Integer
   Dim StyleFlag As Long
   StyleFlag = GetWindowLong(hWndA, GWL_STYLE)
   If StyleFlag And WS_VSCROLL Then
      ScrollBarVisible = True
   Else
      ScrollBarVisible = False
   End If
End Function

' ******************************************************************************
' Routine       : SetCurrentDirectory
' Created by    : Marclei V Silva
' Machine       : ZEUS
' Date-Time     : 01/02/01 2:43:48
' Description   : Set windows current directory
' Inputs        :
' Outputs       :
' Credits       :
' Modifications :
' Remarks       :
' ******************************************************************************
Public Function SetCurrentDirectory(Path As Variant) As Boolean
    Dim strPath As String
    Dim lngRetVal As Long
    
    strPath = Path
    lngRetVal = ApiSetCurrentDirectory(strPath)
    SetCurrentDirectory = (lngRetVal <> 0)
End Function
'-- end code
