VERSION 5.00
Begin VB.UserControl vbalTitleBarModifier 
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   600
   ScaleWidth      =   600
   ToolboxBitmap   =   "vbalTitleBarModifier.ctx":0000
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   60
      Picture         =   "vbalTitleBarModifier.ctx":00FA
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "vbalTitleBarModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_hWnd As Long
Private Const WM_DESTROY = &H2&

Implements ISubclass

Public Sub Attach(ByVal hwnd As Long)
   Detach
   m_hWnd = hwnd
   mTitleBarMod.AttachTitleBarMod hwnd
   AttachMessage Me, m_hWnd, WM_DESTROY
End Sub

Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_DESTROY
      mTitleBarMod.DetachTitleBarMod m_hWnd
      m_hWnd = 0
   End If
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As SSubTimer6.EMsgResponse
   ISubClass_MsgResponse = emrPreprocess
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Select Case iMsg
   Case WM_DESTROY
      Detach
   End Select
End Function

Private Sub UserControl_InitProperties()
   '
End Sub

Private Sub UserControl_Paint()
   UserControl.ForeColor = vb3DHighlight
   UserControl.CurrentX = 0
   UserControl.CurrentY = UserControl.Height
   UserControl.Line -(0, 0)
   UserControl.Line -(UserControl.Width, 0)
   UserControl.ForeColor = vbButtonFace
   UserControl.CurrentX = UserControl.Width - Screen.TwipsPerPixelX
   UserControl.CurrentY = 0
   UserControl.Line -(UserControl.CurrentX, UserControl.Height - Screen.TwipsPerPixelY)
   UserControl.Line -(0, UserControl.Height - Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   '
End Sub

Private Sub UserControl_Resize()
   UserControl.Width = imgIcon.Width + imgIcon.Left * 2
   UserControl.Height = imgIcon.Height + imgIcon.Top * 2
End Sub

Private Sub UserControl_Terminate()
   Detach
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   '
End Sub
