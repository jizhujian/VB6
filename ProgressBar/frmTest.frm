VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form frmTestProgress 
   Caption         =   "Progress Bar Control Tester"
   ClientHeight    =   3645
   ClientLeft      =   5940
   ClientTop       =   4710
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   5730
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar6 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Tag             =   "5"
      Top             =   2700
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   661
      Picture         =   "frmTest.frx":1272
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":2518
      Max             =   250
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar3 
      Height          =   2895
      Left            =   3420
      TabIndex        =   11
      Tag             =   "5"
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5106
      Picture         =   "frmTest.frx":2534
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":2550
      Max             =   293
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar4 
      Height          =   2895
      Left            =   3000
      TabIndex        =   10
      Tag             =   "5"
      Top             =   600
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   5106
      Picture         =   "frmTest.frx":256C
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":2588
      Max             =   500
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar2 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Tag             =   "5"
      Top             =   2040
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      Picture         =   "frmTest.frx":3124
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":3140
      Max             =   150
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar5 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Tag             =   "5"
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      Picture         =   "frmTest.frx":315C
      ForeColor       =   0
      BarColor        =   -2147483635
      BarPicture      =   "frmTest.frx":3178
      Max             =   66
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Tag             =   "5"
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      Picture         =   "frmTest.frx":3194
      ForeColor       =   0
      BarPicture      =   "frmTest.frx":31B0
      BarPictureMode  =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Segments        =   -1  'True
   End
   Begin VB.CommandButton cmdStep 
      Caption         =   "&Step"
      Height          =   435
      Left            =   4440
      TabIndex        =   6
      Top             =   780
      Width           =   1155
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "&Animate"
      Height          =   435
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   1155
   End
   Begin VB.Timer tmrUpd 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3900
      Top             =   3060
   End
   Begin VB.Label lblInfo 
      Caption         =   "Image Processed:"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "Vertical Bars:"
      Height          =   255
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Top             =   300
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Caption         =   "XP:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "System Colour Solid Bar:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      Caption         =   "Stretched Bitmap, Segmented"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   2775
   End
End
Attribute VB_Name = "frmTestProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnimate_Click()
   If tmrUpd.Enabled Then
      tmrUpd.Enabled = False
      cmdStep.Enabled = True
      cmdAnimate.Caption = "&Animate"
   Else
      tmrUpd.Enabled = True
      cmdStep.Enabled = False
      cmdAnimate.Caption = "&Stop"
   End If
End Sub

Private Sub cmdStep_Click()
   tmrUpd_Timer
End Sub

Private Sub Form_Load()
   ' copy bar picture to background picture:
   Set vbalProgressBar6.BarPicture = vbalProgressBar6.Picture
   ' make bar picture lighter and more saturated:
   vbalProgressBar6.ModifyBarPicture 1.7, 2
   ' make background picture less saturated
   vbalProgressBar6.ModifyPicture 1, 0.2
   
   ' copy background picture to bar picture:
   Set vbalProgressBar4.Picture = vbalProgressBar4.BarPicture
   ' make background picture darker and less saturated
   vbalProgressBar4.ModifyPicture 0.6, 0.2
   
End Sub

Private Sub tmrUpd_Timer()
Dim ctl As Control
   For Each ctl In Me.Controls
      If TypeOf ctl Is vbalProgressBar Then
         With ctl
            .Value = .Value + .Tag
            If ctl.ShowText Then
               If ctl.Name = "CSProgressBar4" Then
                  .Text = "Reading: " & .Value & " of " & .Max
               Else
                  .Text = CLng(.Percent) & "%"
               End If
            End If
            If .Value >= .Max Then
               .Tag = -1 * Abs(.Tag)
            ElseIf .Value < 1 Then
               .Tag = Abs(.Tag)
            End If
         End With
      End If
   Next

End Sub
