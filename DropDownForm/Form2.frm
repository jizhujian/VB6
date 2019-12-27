VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.PictureBox picToolbar 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   4680
      TabIndex        =   0
      Top             =   0
      Width           =   4680
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
         Left            =   3180
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   60
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
         TabIndex        =   3
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
End
Attribute VB_Name = "Form2"
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

Private Const SW_SHOWNOACTIVATE = 4
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

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

Private Sub chkDropDown_Click()
   Position DropDownWIndow, txtDropDown
   ShowWindow DropDownWIndow.hwnd, SW_SHOWNOACTIVATE
  SetWindowPos DropDownWIndow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
   With DropDownWIndow
      .ShowState = ewssDropped
      .Value = txtDropDown.Text
   End With
End Sub

Private Sub cmdPalette_Click()
   Position PaletteWindow, cmdPalette
   ShowWindow PaletteWindow.hwnd, SW_SHOWNOACTIVATE
  SetWindowPos PaletteWindow.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
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

Private Sub Form_Load()
   
   chkDropDown.Move txtDropDown.Left + txtDropDown.Width - chkDropDown.Width - 2 * Screen.TwipsPerPixelX, txtDropDown.Top + 2 * Screen.TwipsPerPixelY, chkDropDown.Width, txtDropDown.Height - 4 * Screen.TwipsPerPixelY
   
End Sub

