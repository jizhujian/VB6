VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim f As New frmProgress

Private Const SW_SHOWNOACTIVATE = 4
Private Const HWND_TOPMOST = -1
Private Const HWND_TOP = 0
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Sub Show(Optional ByVal topMost As Boolean)
  Dim parentHWnd As Long
  Dim parentRect As RECT
  Dim meLeft As Long
  Dim meTop As Long
  parentHWnd = GetActiveWindow()
  GetClientRect parentHWnd, parentRect
  meLeft = ((parentRect.Right - parentRect.Left) * Screen.TwipsPerPixelX - f.Width) / 2
  meTop = ((parentRect.Bottom - parentRect.Top) * Screen.TwipsPerPixelY - f.Height) / 2
  If (meLeft + f.Width > Screen.Width) Then meLeft = Screen.Width - f.Width
  If (meTop + f.Height > Screen.Height) Then meTop = Screen.Height - f.Height
  If (meLeft < 0) Then meLeft = 0
  If (meTop < 0) Then meTop = 0
  f.Move meLeft, meTop
  SetParent f.hwnd, parentHWnd
  ShowWindow f.hwnd, SW_SHOWNOACTIVATE
  SetWindowPos f.hwnd, IIf(topMost, HWND_TOPMOST, HWND_TOP), 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
End Sub

Public Property Get Text() As String
  Text = f.pb.Text
End Property

Public Property Let Text(ByVal s As String)
  f.pb.Text = s
End Property

Public Property Get Value() As Long
  Value = f.pb.Value
End Property

Public Property Let Value(ByVal v As Long)
  f.pb.Value = v
End Property

Public Property Get Min() As Long
  Min = f.pb.Min
End Property

Public Property Let Min(ByVal l As Long)
  f.pb.Min = l
End Property

Public Property Get Max() As Long
  Max = f.pb.Max
End Property

Public Property Let Max(ByVal u As Long)
  f.pb.Max = u
End Property

Private Sub Class_Terminate()
  Unload f
  Set f = Nothing
End Sub
