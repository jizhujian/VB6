VERSION 5.00
Begin VB.Form frmInputString 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入字符串"
   ClientHeight    =   1545
   ClientLeft      =   30
   ClientTop       =   420
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtValue 
      Height          =   264
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4692
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   372
      Left            =   3840
      TabIndex        =   3
      Top             =   600
      Width           =   972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   372
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   972
   End
   Begin VB.Label lblPrompt 
      Caption         =   "请输入字符串："
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputString"
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
