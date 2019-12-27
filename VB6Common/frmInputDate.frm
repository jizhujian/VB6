VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmInputDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入日期"
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
      TabIndex        =   1
      Top             =   600
      Width           =   972
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4692
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
      DateIsNull      =   -1  'True
      Format          =   23199747
      CurrentDate     =   41165
   End
   Begin VB.Label lblPrompt 
      Caption         =   "请输入日期："
      Height          =   972
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3492
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mblnSuccess As Boolean
Public mblnCheckMinValue As Boolean
Public mdtmMinValue As Date
Public mblnCheckMaxValue As Boolean
Public mdtmMaxValue As Date

Private Sub cmdOK_Click()
  If Not IsNull(dtpDate.Value) Then
    If mblnCheckMinValue Then
      If dtpDate.Value < mdtmMinValue Then
        MsgBox "不能小于最小值 ：" & Format(mdtmMinValue, dtpDate.CustomFormat), vbInformation, Caption
        dtpDate.SetFocus
        Exit Sub
      End If
    End If
    If mblnCheckMaxValue Then
      If dtpDate.Value > mdtmMaxValue Then
        MsgBox "不能大于最大值 ：" & Format(mdtmMaxValue, dtpDate.CustomFormat), vbInformation, Caption
        dtpDate.SetFocus
        Exit Sub
      End If
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
