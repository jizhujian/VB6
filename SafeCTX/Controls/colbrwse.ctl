VERSION 5.00
Begin VB.UserControl ColorBrowser 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   KeyPreview      =   -1  'True
   ScaleHeight     =   360
   ScaleWidth      =   2535
   ToolboxBitmap   =   "colbrwse.ctx":0000
   Begin VB.CommandButton cmdBrowse 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1815
      Picture         =   "colbrwse.ctx":0312
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   270
   End
   Begin VB.PictureBox picSelection 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   30
      ScaleHeight     =   240
      ScaleWidth      =   1755
      TabIndex        =   0
      Top             =   15
      Width           =   1755
   End
End
Attribute VB_Name = "ColorBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type ChooseColor
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long


'Event Declarations:
Event Change(NewColor As Long)

'Default Property Values:
Const m_def_SelectedColor = vbWhite
Const m_def_AllowCustomColors = True

'Property Variables:
Dim m_AllowCustomColors As Boolean
Dim m_SelectedColor As Long

Public Sub AboutBox()
    About
End Sub

Private Sub cmdBrowse_Click()
    picSelection.SetFocus 'so we dont see the Focus Rectangle
    'Show the Color Dialog
    If GetColor(m_SelectedColor, m_AllowCustomColors) Then
        picSelection_Paint
        RaiseEvent Change(m_SelectedColor)
    End If
End Sub

Private Sub picSelection_Click()
    'Fire the click event
    cmdBrowse_Click
End Sub

Private Sub picSelection_GotFocus()
    picSelection_Paint
End Sub

Private Sub picSelection_LostFocus()
    picSelection_Paint
End Sub

Private Sub picSelection_Paint()
    'Draw a focus rectangle
    Dim rct As RECT
    
    If GetFocus = picSelection.hWnd Then
        GetClientRect picSelection.hWnd, rct
        With rct
            .Left = .Left + 1
            .Right = .Right - 1
            .Top = .Top + 1
            .Bottom = .Bottom - 1
        End With
        DrawFocusRect picSelection.hdc, rct
    Else
        picSelection.Cls
    End If
    'Paint the interior with the selected color
    picSelection.Line (2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY)-(picSelection.ScaleWidth - (3 * Screen.TwipsPerPixelX), picSelection.ScaleHeight - (3 * Screen.TwipsPerPixelY)), m_SelectedColor, BF
    
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the Color dialog to show
    'Alt and down arrow
    If KeyCode = vbKeyDown And Shift = 4 Then
        cmdBrowse_Click
    End If
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'Position the constituent controls
    cmdBrowse.Move UserControl.ScaleWidth - cmdBrowse.Width, 0, cmdBrowse.Width, UserControl.ScaleHeight
    picSelection.Move 0, 0, UserControl.ScaleWidth - (cmdBrowse.Width + Screen.TwipsPerPixelX), UserControl.ScaleHeight
End Sub

Private Sub UserControl_Show()
    'Get the tooltip
    picSelection.ToolTipText = UserControl.Extender.ToolTipText
End Sub

Public Property Let SelectedColor(New_SelectedColor As Long)
    m_SelectedColor = New_SelectedColor
    picSelection_Paint
End Property

Public Property Get SelectedColor() As Long
    SelectedColor = m_SelectedColor
End Property

Private Function GetColor(lColor As Long, Optional bAllowCustomColors As Boolean = False) As Boolean
    
    Dim cc As ChooseColor

    With cc
        .lStructSize = Len(cc)
        .hWndOwner = UserControl.hWnd
        .hInstance = App.hInstance
        If Not bAllowCustomColors Then
            .flags = &H4
        End If
        .lpCustColors = String$(16 * 4, 0)
    End With
    
    If ChooseColor(cc) > 0 Then
        lColor = cc.rgbResult
        GetColor = True
    End If

End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AllowCustomColors() As Boolean
    AllowCustomColors = m_AllowCustomColors
End Property

Public Property Let AllowCustomColors(ByVal New_AllowCustomColors As Boolean)
    m_AllowCustomColors = New_AllowCustomColors
    PropertyChanged "AllowCustomColors"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowCustomColors = m_def_AllowCustomColors
    m_SelectedColor = m_def_SelectedColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_AllowCustomColors = PropBag.ReadProperty("AllowCustomColors", m_def_AllowCustomColors)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AllowCustomColors", m_AllowCustomColors, m_def_AllowCustomColors)
End Sub

