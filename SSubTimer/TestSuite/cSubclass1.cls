VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSubclass1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MOUSEMOVE = &H200
Private Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5

Implements ISubclass

Private m_hWnd As Long

Public Sub Attach(ByVal hWnd As Long)
   Detach
   attachmessage Me, hWnd, WM_LBUTTONUP
   attachmessage Me, hWnd, WM_SIZE
   attachmessage Me, hWnd, WM_DESTROY
   m_hWnd = hWnd
End Sub

Public Sub Detach()
   If Not (m_hWnd = 0) Then
      detachmessage Me, m_hWnd, WM_LBUTTONUP
      detachmessage Me, m_hWnd, WM_SIZE
      detachmessage Me, m_hWnd, WM_DESTROY
      m_hWnd = 0
   End If
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   '
   Select Case CurrentMessage
   Case WM_LBUTTONUP
      ISubclass_MsgResponse = emrPostProcess
   Case WM_SIZE
      ISubclass_MsgResponse = emrPostProcess
   Case WM_DESTROY
      ISubclass_MsgResponse = emrPostProcess
   End Select
   '
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   '
   Select Case CurrentMessage
   Case WM_LBUTTONUP
      'Debug.Print "Class1:WM_LBUTTONUP"
   Case WM_SIZE
      'Debug.Print "Class1:WM_SIZE"
   Case WM_DESTROY
      'Debug.Print "Class1:WM_DESTROY"
      Detach
   End Select
   '
End Function
