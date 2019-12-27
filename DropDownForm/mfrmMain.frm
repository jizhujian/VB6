VERSION 5.00
Object = "{4A3A29A4-F2E3-11D3-B06C-00500427A693}#4.0#0"; "vbalDDFm6.ocx"
Begin VB.MDIForm mfrmTest 
   BackColor       =   &H8000000C&
   Caption         =   "vbAccelerator Floating Tool Window Tester"
   ClientHeight    =   5205
   ClientLeft      =   2190
   ClientTop       =   2160
   ClientWidth     =   7335
   Icon            =   "mfrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin vbalDropDownForm6.vbalTitleBarModifier ctlTitleBarMod 
      Left            =   60
      Top             =   600
      _ExtentX        =   635
      _ExtentY        =   635
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.CheckBox chkDropDown 
         Caption         =   "u"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   8.25
            Charset         =   2
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         Width           =   195
      End
      Begin VB.CommandButton cmdPalette 
         Caption         =   "Palette"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   60
         Width           =   1035
      End
      Begin VB.TextBox txtDropDown 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   60
         Width           =   2295
      End
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Find..."
         Index           =   0
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Find &Next"
         Index           =   1
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Replace..."
         Index           =   2
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Palette..."
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&vbAccelerator on the Web..."
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   2
      End
   End
End
Attribute VB_Name = "mfrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private WithEvents m_fDropDown As frmDropDown
Attribute m_fDropDown.VB_VarHelpID = -1
Private WithEvents m_fFindReplace As frmFindReplace
Attribute m_fFindReplace.VB_VarHelpID = -1
Private m_fPalette As frmPalette
Attribute m_fPalette.VB_VarHelpID = -1

Private Const SW_NORMAL = 1
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Property Get PaletteWindow() As frmPalette
   If m_fPalette Is Nothing Then
      Set m_fPalette = New frmPalette
   End If
   Set PaletteWindow = m_fPalette
End Property

Private Property Get DropDownWIndow() As frmDropDown
   If m_fDropDown Is Nothing Then
      Set m_fDropDown = New frmDropDown
   End If
   Set DropDownWIndow = m_fDropDown
End Property
Private Property Get FindReplaceWIndow() As frmFindReplace
   If m_fFindReplace Is Nothing Then
      Set m_fFindReplace = New frmFindReplace
   End If
   Set FindReplaceWIndow = m_fFindReplace
End Property

Private Sub Position(frmThis As Form, objThis As Object)
Dim tR As RECT
   GetWindowRect objThis.hwnd, tR
   frmThis.Move tR.Left * Screen.TwipsPerPixelX, (tR.Bottom + 1) * Screen.TwipsPerPixelY
End Sub
Public Sub ShowFindReplace(ByVal bReplaceMode As Boolean)
   With FindReplaceWIndow
      .Show , Me
      .ShowState = ewssFloating
      If bReplaceMode Then
         .Mode = efrReplace
      Else
         .Mode = efrFind
      End If
   End With
End Sub

Private Sub chkDropDown_Click()
   Position DropDownWIndow, txtDropDown
   ShowWindow DropDownWIndow.hwnd, SW_NORMAL
   With DropDownWIndow
'      .Show , Me
      .ShowState = ewssDropped
      .Value = txtDropDown.Text
   End With
End Sub

Private Sub cmdPalette_Click()
   Position PaletteWindow, cmdPalette
   ShowWindow PaletteWindow.hwnd, SW_NORMAL
   With PaletteWindow
      ' Show the window & set the drop-state
'      .Show , Me
      .ShowState = ewssDropped
   End With
End Sub

Private Sub m_fDropDown_Change()
   txtDropDown.Text = DropDownWIndow.Value
End Sub

Private Sub m_fDropDown_CloseUp()
   If Not DropDownWIndow.Cancelled Then
      txtDropDown.Text = DropDownWIndow.Value
      txtDropDown.SelLength = Len(txtDropDown.Text)
   End If
   chkDropDown.Value = Unchecked
   txtDropDown.SetFocus
End Sub

Private Sub MDIForm_Load()

   mnuFile_Click 0
   
   chkDropDown.Move txtDropDown.Left + txtDropDown.Width - chkDropDown.Width - 2 * Screen.TwipsPerPixelX, txtDropDown.Top + 2 * Screen.TwipsPerPixelY, chkDropDown.Width, txtDropDown.Height - 4 * Screen.TwipsPerPixelY
   
   ctlTitleBarMod.Attach Me.hwnd
   
End Sub

Private Sub mnuEdit_Click(Index As Integer)
   
   Select Case Index
   Case 0 'Find
      ShowFindReplace False
   Case 1 'F N
      
   Case 2 'Replace
      ShowFindReplace True
   Case 4 'Palette
   
   End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
   Select Case Index
   Case 0
      Dim f As New frmMDIChild
      f.Show
   Case 2
      Unload Me
   End Select
End Sub

Private Sub mnuHelp_Click(Index As Integer)
   Select Case Index
   Case 0
      ShellEx "http://vbaccelerator.com/"
   Case 2
      frmAbout.Show vbModal, Me
   End Select
End Sub

Private Sub txtDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyF4
      chkDropDown_Click
      chkDropDown.Value = Checked
      KeyCode = 0
   Case vbKeyDown
      ' Like extended UI in ComboBox
      chkDropDown_Click
      chkDropDown.Value = Checked
      KeyCode = 0
   Case Else
      ' .
   End Select
End Sub

