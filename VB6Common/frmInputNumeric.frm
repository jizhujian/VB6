VERSION 5.00
Begin VB.Form frmInputNumeric 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入数值"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   372
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   972
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
   Begin VB.TextBox txtValue 
      Height          =   264
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   4692
   End
   Begin VB.Label lblPrompt 
      Caption         =   "请输入数值："
      Height          =   972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3492
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnSuccess As Boolean
Public mblnCheckMinValue As Boolean
Public mdblMinValue As Double
Public mblnCheckMaxValue As Boolean
Public mdblMaxValue As Double

Private Sub cmdOK_Click()
  Dim dblValue As Double
  dblValue = Val(txtValue.Text)
  If mblnCheckMinValue Then
    If dblValue < mdblMinValue Then
      MsgBox "不能小于最小值 ：" & mdblMinValue, vbInformation, Caption
      txtValue.SetFocus
      Exit Sub
    End If
  End If
  If mblnCheckMaxValue Then
    If dblValue > mdblMaxValue Then
      MsgBox "不能大于最大值 ：" & mdblMaxValue, vbInformation, Caption
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
