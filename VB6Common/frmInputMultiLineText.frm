VERSION 5.00
Begin VB.Form frmInputMultiLineText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入多行文本"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtValue 
      Height          =   3375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6615
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "请输入多行文本："
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "frmInputMultiLineText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnSuccess As Boolean
Public mblnAllowEmpty As Boolean

Private Sub cmdOK_Click()
  If Not mblnAllowEmpty Then
    If (Trim(txtValue.Text) = "") Then
      MsgBox "必须输入值。", vbInformation, Caption
      txtValue.SetFocus
      Exit Sub
    End If
  End If
  mblnSuccess = True
  Hide
End Sub

Private Sub cmdCancel_Click()
  Hide
End Sub

Private Sub Form_Load()
  Dim res As New IconResource
  Set Icon = res.LoadResIcon("color_swatch")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Hide
  End If
End Sub

Private Sub txtValue_GotFocus()
  With txtValue
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

