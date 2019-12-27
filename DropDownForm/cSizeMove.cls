VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSizeMoveHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===================================================================
' cSizeMoveHelper.cls
' Author:   Steve McMahon
' Date:     10 August 1999
'
' Allows a form to take advanced control of the moving
' and sizing process, including
'
' -------------------------------------------------------------------
' vbAccelerator - Advanced, Free Source Code for Visual Basic
' http://vbaccelerator.com/
' mailto:steve@vbaccelerator.com
' ===================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Const WM_DESTROY = &H2
Private Const WM_SIZING = &H214
Private Const WM_MOVING = &H216&
Private Const WM_ENTERSIZEMOVE = &H231&
Private Const WM_EXITSIZEMOVE = &H232&
Private Const WM_ACTIVATE = &H6
Private Const WM_SIZE = &H5

Private m_hWnd As Long
Private m_bActive As Boolean
Private m_bInSizeMove As Boolean

Implements ISubclass

Public Event EnterSizeMove()
Public Event ExitSizeMove()
Public Event Activate(ByVal bByMouse As Boolean)
Public Event Deactivate()
Public Event Moving(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Public Event Sizing(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)

Public Property Get Active() As Boolean
   Active = m_bActive
End Property

Public Sub Attach(ByVal hWndA As Long)
   Detach
   m_hWnd = hWndA
   AttachMessage Me, m_hWnd, WM_ENTERSIZEMOVE
   AttachMessage Me, m_hWnd, WM_EXITSIZEMOVE
   AttachMessage Me, m_hWnd, WM_DESTROY
   AttachMessage Me, m_hWnd, WM_MOVING
   AttachMessage Me, m_hWnd, WM_SIZING
   AttachMessage Me, m_hWnd, WM_ACTIVATE
   AttachMessage Me, m_hWnd, WM_SIZE
End Sub
Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_ENTERSIZEMOVE
      DetachMessage Me, m_hWnd, WM_EXITSIZEMOVE
      DetachMessage Me, m_hWnd, WM_DESTROY
      DetachMessage Me, m_hWnd, WM_MOVING
      DetachMessage Me, m_hWnd, WM_SIZING
      DetachMessage Me, m_hWnd, WM_SIZE
      DetachMessage Me, m_hWnd, WM_ACTIVATE
   End If
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As SSubTimer6.EMsgResponse
   ' Process before windows:
   ISubClass_MsgResponse = emrPostProcess
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lR As Long
Dim tR As RECT
Dim lWidth As Long
Dim lHeight As Long

   Select Case iMsg
   Case WM_ENTERSIZEMOVE
      ' Enter modal sizing/moving loop
      m_bInSizeMove = True
      RaiseEvent EnterSizeMove
      
   Case WM_EXITSIZEMOVE
      ' Exit modal sizing/moving loop
      RaiseEvent ExitSizeMove
      m_bInSizeMove = False
      
   Case WM_MOVING
      ' Form is moving:
      CopyMemory tR, ByVal lParam, Len(tR)
      lWidth = tR.Right - tR.Left
      lHeight = tR.Bottom - tR.Top
      RaiseEvent Moving(tR.Left, tR.Top, lWidth, lHeight)
      tR.Right = tR.Left + lWidth
      tR.Bottom = tR.Top + lHeight
      CopyMemory ByVal lParam, tR, Len(tR)
      
   Case WM_SIZING
      ' Form is sizing:
      CopyMemory tR, ByVal lParam, Len(tR)
      lWidth = tR.Right - tR.Left
      lHeight = tR.Bottom - tR.Top
      RaiseEvent Sizing(tR.Left, tR.Top, lWidth, lHeight)
      tR.Right = tR.Left + lWidth
      tR.Bottom = tR.Top + lHeight
      CopyMemory ByVal lParam, tR, Len(tR)
      
   Case WM_ACTIVATE
      ' Form is activated/deactivated:
      If wParam = 0 Then
         ' deactivate:
         m_bActive = False
         RaiseEvent Deactivate
      Else
         ' Activate
         m_bActive = True
         RaiseEvent Activate(wParam = 2)
      End If

   'Case WM_SIZE
   '   ' This ensures that you don't have to separately
   '   ' check for maximize/minimize/restore/in code
   '   ' movement, as these aren't
   '   ' controlled by the UI sizing loop:
   '   If Not (m_bInSizeMove) Then
   '      GetWindowRect m_hWnd, tR
   '      lWidth = tR.Right - tR.Left
   '      lHeight = tR.Bottom - tR.Top
   '      RaiseEvent Sizing(tR.Left, tR.Top, lWidth, lHeight)
   '   End If
      
   Case WM_DESTROY
      ' Ensures the class terminates regardless
      ' of whether the user explicitly detaches
      ' it or not:
      'Debug.Print "WM_DESTROY"
      Detach
   End Select
End Function
