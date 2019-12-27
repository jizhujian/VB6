VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbAccelerator's Goldfish"
   ClientHeight    =   5175
   ClientLeft      =   5880
   ClientTop       =   3450
   ClientWidth     =   5550
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3571.877
   ScaleMode       =   0  'User
   ScaleWidth      =   5211.736
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4140
      TabIndex        =   2
      Top             =   4620
      Width           =   1260
   End
   Begin VB.TextBox txtSpecialCopyright 
      Height          =   675
      Left            =   1140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "fAbout.frx":0000
      Top             =   5640
      Width           =   4455
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Form Control"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1620
      TabIndex        =   13
      Top             =   1020
      Width           =   3375
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00C0C000&
      BorderWidth     =   16
      Height          =   855
      Index           =   1
      Left            =   180
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "vbAccelerator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   480
      Left            =   1140
      TabIndex        =   8
      Top             =   180
      Width           =   1725
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.00"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      TabIndex        =   12
      Top             =   3960
      Width           =   1410
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   957.833
      X2              =   5253.993
      Y1              =   2691.849
      Y2              =   2691.849
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   360
      Picture         =   "fAbout.frx":00C8
      Top             =   420
      Width           =   480
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      Height          =   675
      Index           =   0
      Left            =   4500
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Drop-Down"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   1620
      TabIndex        =   9
      Top             =   480
      Width           =   2835
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 1999 Steve McMahon "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "fAbout.frx":133A
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Tag             =   "http://vbaccelerator.com/j-index.htm?url=cright.htm"
      Top             =   4620
      Width           =   3990
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://vbaccelerator.com/"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "fAbout.frx":1644
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "http://vbaccelerator.com/"
      Top             =   4800
      Width           =   3870
   End
   Begin VB.Label lblDotLine 
      BackStyle       =   0  'Transparent
      Caption         =   "................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2880
      TabIndex        =   4
      Top             =   60
      Width           =   2955
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "vbAccelerator Drop-Down Form Control  Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1080
      MouseIcon       =   "fAbout.frx":194E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "http://vbaccelerator.com/j-index.htm?url=codelib/ddtoolwn/ddform.htm"
      Top             =   4140
      Width           =   3855
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"fAbout.frx":1C58
      ForeColor       =   &H00000000&
      Height          =   810
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   1860
      Width           =   3885
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H0080C0FF&
      Height          =   225
      Index           =   4
      Left            =   1080
      TabIndex        =   11
      Top             =   540
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H00000000&
      Height          =   870
      Index           =   3
      Left            =   1080
      TabIndex        =   10
      Top             =   780
      Width           =   3885
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0E0FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   4860
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   Me.Caption = "About " & App.Title
   Me.Icon = mfrmTest.Icon
   lblVersion.Caption = "Version: " & App.Major & "." & Format$(App.Minor, "00") & "." & Format$(App.Revision, "0000")
End Sub

Private Sub imgVBAccelerator_Click()
   pShell "http://vbaccelerator.com/"
End Sub

Private Sub pShell(ByVal sWhat As String)
   On Error Resume Next
   ShellEx sWhat, , , , , Me.hwnd
   If (Err.Number <> 0) Then
       MsgBox "Sorry, I failed to open '" & sWhat & "' due to an error." & vbCrLf & vbCrLf & "[" & Err.Description & "]", vbExclamation
   End If
End Sub

Private Sub lblCopyright_Click()
   pShell lblCopyright.Tag
End Sub

Private Sub lblProduct_Click()
   pShell lblProduct.Tag
End Sub

Private Sub lblURL_Click()
   pShell lblURL.Tag
End Sub
