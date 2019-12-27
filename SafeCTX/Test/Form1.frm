VERSION 5.00
Object = "{2A544CFD-5E25-11D3-8002-00A0C93E2B7E}#2.0#0"; "SafeCTX.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2940
   ClientLeft      =   6825
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin SafeCTX.TrayArea TrayArea1 
      Left            =   4200
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
   End
   Begin SafeCTX.Label3D Label3D1 
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483634
      ForeColor2      =   -2147483630
      Caption         =   "FailSafe Systems"
      Alignment       =   2
      BorderStyle     =   1
   End
   Begin SafeCTX.Splitter Splitter1 
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   3625
      SplitWidth      =   50
      Begin VB.Frame Frame2 
         Caption         =   "Colors"
         Height          =   1815
         Left            =   2040
         TabIndex        =   6
         Top             =   0
         Width           =   1815
         Begin SafeCTX.ColorSelector ColorSelector1 
            Height          =   315
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
         End
         Begin SafeCTX.ColorButton ColorButton1 
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
         End
         Begin SafeCTX.ColorBrowser ColorBrowser1 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Browse"
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1815
         Begin SafeCTX.FontBrowser FontBrowser1 
            Height          =   315
            Left            =   120
            TabIndex        =   2
            Top             =   960
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            ForeColor       =   -2147483630
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SafeCTX.FileBrowser FileBrowser1 
            Height          =   315
            Left            =   120
            TabIndex        =   3
            Top             =   600
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SafeCTX.FolderBrowser FolderBrowser1 
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SafeCTX.Hyperlink Hyperlink1 
            Height          =   315
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   1
            Hyperlink       =   "http://www.failsafe.co.za"
            Caption         =   "Visit our WebSite"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    TrayArea1.Visible = True
    Set TrayArea1.Icon = Me.Icon

    'Splitter Control
    'Splitter1.Left = 0
    Splitter1.SplitPercent = 50
    Set Splitter1.Control1 = Frame1
    Set Splitter1.Control2 = Frame2
    'You can of course have another splitter
    'as the second control and split again eg:
    'Set Splitter1.Control1 = tvMain
    'Set Splitter1.Control2 = Splitter2
    'Set Splitter2.Control1 = lvMain
    'Set Splitter2.Control2 = Splitter3
End Sub

Private Sub Form_Paint()
    'Form_Resize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        With Splitter1
            .Top = Me.ScaleTop + Label3D1.Height
            .Left = Me.ScaleLeft
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - Label3D1.Height
        End With
    End If
    
    Dim oCtl As Control
    Label3D1.Width = Me.ScaleWidth
    For Each oCtl In Me.Controls
        Select Case oCtl.Name
            Case "FolderBrowser1", "FileBrowser1", "FontBrowser1", "Hyperlink1"
                oCtl.Width = Frame1.Width - 250
            Case "ColorBrowser1", "ColorButton1", "ColorSelector1"
                oCtl.Width = Frame2.Width - 250
        End Select
    Next
End Sub

Private Sub Splitter1_GotFocus()
    '
End Sub

Private Sub Splitter1_LostFocus()
    '
End Sub

Private Sub Splitter1_Validate(Cancel As Boolean)
    Form_Resize
End Sub

Private Sub TrayArea1_MouseMove()
    If Me.WindowState = vbMinimized Then
        TrayArea1.ToolTip = "Click to Restore " & App.EXEName
    Else
        TrayArea1.ToolTip = "Click to Minimize " & App.EXEName
    End If
End Sub

Private Sub TrayArea1_MouseUp(Button As Integer)
    With Me
    If .WindowState = vbMinimized Then
        .WindowState = vbNormal
        .Visible = True
        .SetFocus
    Else
        .WindowState = vbMinimized
        .Visible = False
    End If
    End With
End Sub
