VERSION 5.00
Begin VB.Form frmSSubTmr 
   Caption         =   "vbAccelerator SSubTmr Tester"
   ClientHeight    =   2925
   ClientLeft      =   3870
   ClientTop       =   2850
   ClientWidth     =   5085
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   5085
   Begin VB.ListBox lstMsg 
      Height          =   645
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4875
   End
   Begin VB.CheckBox chkTimer2 
      Caption         =   "Timer 2"
      Height          =   375
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Value           =   1  'Checked
      Width           =   1155
   End
   Begin VB.CheckBox chkTimer1 
      Caption         =   "Timer 1"
      Height          =   255
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.Label lblSubClass 
      Caption         =   $"frmTest.frx":014A
      Height          =   735
      Left            =   60
      TabIndex        =   5
      Top             =   1260
      Width           =   4995
   End
   Begin VB.Label lblTimer 
      Caption         =   "All Code Timer Test:"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   4995
   End
   Begin VB.Label lblTimer2 
      Height          =   255
      Left            =   1500
      TabIndex        =   3
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label lblTimer1 
      Height          =   255
      Left            =   1500
      TabIndex        =   2
      Top             =   300
      Width           =   1035
   End
End
Attribute VB_Name = "frmSSubTmr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_t1 As CTimer
Attribute m_t1.VB_VarHelpID = -1
Private WithEvents m_t2 As CTimer
Attribute m_t2.VB_VarHelpID = -1

Implements ISubClass
Private m_emr As EMsgResponse

Private Const WM_SIZE = &H5
Private Const WM_LBUTTONDOWN = &H201

Private Sub ShowMessage(ByVal msg As String)
   Debug.Print msg
   lstMsg.AddItem msg
   lstMsg.ListIndex = lstMsg.ListCount - 1
End Sub

Private Sub chkTimer1_Click()
Dim lI As Long
    If (chkTimer1.Value <> 0) Then
        lI = 100
    Else
        lI = -1
    End If
    m_t1.Interval = lI
End Sub

Private Sub chkTimer2_Click()
Dim lI As Long
    If (chkTimer2.Value <> 0) Then
        lI = 100
    Else
        lI = -1
    End If
    m_t2.Interval = lI
End Sub

Private Sub Form_Load()
    AttachMessage Me, Me.hWnd, WM_LBUTTONDOWN
    AttachMessage Me, Me.hWnd, WM_SIZE
    
    Set m_t1 = New CTimer
    m_t1.Interval = 100
    Set m_t2 = New CTimer
    m_t2.Interval = 100
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DetachMessage Me, Me.hWnd, WM_LBUTTONDOWN
    DetachMessage Me, Me.hWnd, WM_SIZE
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   lstMsg.Move lstMsg.Left, lstMsg.Top, Me.ScaleWidth - lstMsg.Left * 2, Me.ScaleHeight - lstMsg.Top - lstMsg.Left
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As EMsgResponse
   ShowMessage "ISubClass_MsgResponse: CurrentMessage=" & CurrentMessage
   ISubClass_MsgResponse = emrPostProcess
End Property

Private Function ISubClass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ShowMessage "ISubClass_WindowProc: hWnd=" & hWnd & ", iMsg=" & iMsg & ", wParam=" & wParam & ", lParam=" & lParam
End Function

Private Sub m_t1_ThatTime()
    lblTimer1.Caption = Format$(Now, "hh:nn:ss")
End Sub

Private Sub m_t2_ThatTime()
    lblTimer2.Caption = Format$(Now, "hh:nn:ss")
End Sub
