VERSION 5.00
Object = "{4A3A29A4-F2E3-11D3-B06C-00500427A693}#2.0#0"; "vbalDDFm6.ocx"
Begin VB.Form frmPalette 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1965
   ClientLeft      =   3735
   ClientTop       =   4005
   ClientWidth     =   4095
   ControlBox      =   0   'False
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
   ScaleHeight     =   1965
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin vbalDropDownForm6.vbalDropDownClient ddcDropDOwn 
      Align           =   1  'Align Top
      Height          =   75
      Left            =   0
      ToolTipText     =   "Drag to make this menu float"
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   132
      Caption         =   "Palette"
   End
   Begin VB.PictureBox picControlHolder 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   4035
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4035
      Begin VB.PictureBox picMore 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   675
         Left            =   720
         ScaleHeight     =   45
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   5
         Top             =   1200
         Width           =   2355
      End
      Begin VB.PictureBox picStandard 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   720
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   157
         TabIndex        =   4
         Top             =   360
         Width           =   2355
      End
      Begin VB.CheckBox chkMoreColours 
         Caption         =   "&More Colours"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   2355
      End
      Begin VB.TextBox txtColour 
         Height          =   315
         Left            =   720
         TabIndex        =   0
         Text            =   "&H000000"
         Top             =   60
         Width           =   2355
      End
      Begin VB.Label lblColour 
         Caption         =   "&Colour:"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bReplaceMode As Boolean
Private m_lS As Long

Public Property Get ReplaceMode() As Boolean
   ReplaceMode = m_bReplaceMode
End Property
Public Property Let ReplaceMode(ByVal bState As Boolean)
   m_bReplaceMode = bState
End Property

Public Property Let ShowState(ByVal eState As EWindowShowState)
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ddcDropDown.ShowState = eState
End Property
Public Property Get ShowState() As EWindowShowState
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ShowState = ddcDropDown.ShowState
End Property

Private Sub chkMoreColours_Click()
   If chkMoreColours.Value = Checked Then
      pRenderMoreColours
   Else
      picControlHolder.Height = picMore.Top
      ddcDropDown_CaptionResize
   End If
End Sub

Private Sub ddcDropDown_AppActivate(ByVal bState As Boolean)
   ' Emulate Word - hide away floating
   ' toolwindows when we're not the focus
   ' app:
'   If (bState) Then
'      If ddcDropDown.ShowState = ewssFloating Then
'         On Error Resume Next ' we might be showing a modal form
'         Me.Show
'      End If
'   Else
'      If ddcDropDown.ShowState = ewssFloating Then
'         On Error Resume Next ' we might be showing a modal form
'         Me.Hide
'      End If
'   End If
End Sub

Private Sub ddcDropDown_CaptionResize()
   ' Here you would resize your form/move the
   ' contents to accommodate the change in size
   ' of the caption:
   If Not ddcDropDown.Sizing Then
      Me.Height = ddcDropDown.Height + 2 * Screen.TwipsPerPixelY + picControlHolder.Height + (Me.Height - Me.ScaleHeight)
      picControlHolder.Top = ddcDropDown.Height + 2 * Screen.TwipsPerPixelY
   End If
End Sub

Private Sub ddcDropDown_Sizing(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
   ' Set the drop-down size within limits:
   If lWidth > 400 Then
      lWidth = 400
   End If
   If lWidth < 96 Then
      lWidth = 96
   End If
   If lHeight > 300 Then
      lHeight = 300
   End If
   If lHeight < 96 Then
      lHeight = 96
   End If
End Sub

Private Sub Form_Load()
   '
   pRenderStandardColours
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
End Sub

Private Sub pRenderStandardColours()
Dim X As Long
Dim Y As Long
Dim i As Long
   m_lS = picStandard.ScaleWidth \ 10
   X = 0
   Y = 0
   For i = 1 To 16
      picStandard.Line (X, Y)-(X + m_lS - 2, Y + m_lS - 2), pOfficeColour(False, i), BF
      picStandard.Line (X, Y)-(X + m_lS - 2, Y + m_lS - 2), vbButtonShadow, B
      X = X + m_lS
      If i = 8 Then
         X = 0
         Y = Y + m_lS
      End If
   Next i
   picStandard.Refresh
End Sub
Private Sub pRenderMoreColours()
Dim X As Long
Dim Y As Long
Dim i As Long
   picMore.Height = (m_lS * 5) * Screen.TwipsPerPixelY
   X = 0
   Y = 0
   For i = 1 To 40
      picMore.Line (X, Y)-(X + m_lS - 2, Y + m_lS - 2), pOfficeColour(True, i), BF
      picMore.Line (X, Y)-(X + m_lS - 2, Y + m_lS - 2), vbButtonShadow, B
      X = X + m_lS
      If i Mod 8 = 0 Then
         X = 0
         Y = Y + m_lS
      End If
   Next i
   picMore.Refresh
   picControlHolder.Height = picMore.Top + picMore.Height
   ddcDropDown_CaptionResize
End Sub
Private Function pOfficeColour(ByVal bLargePalette As Boolean, ByVal nIndex As Long) As OLE_COLOR
   If bLargePalette Then
      Select Case nIndex
      Case 1: pOfficeColour = &H0&
      Case 2: pOfficeColour = &H3399&
      Case 3: pOfficeColour = &H3333&
      Case 4: pOfficeColour = &H3300&
      Case 5: pOfficeColour = &H663300
      Case 6: pOfficeColour = &H800000
      Case 7: pOfficeColour = &H993333
      Case 8: pOfficeColour = &H333333
      
      Case 9: pOfficeColour = &H80&
      Case 10: pOfficeColour = &H66FF&
      Case 11: pOfficeColour = &H8080&
      Case 12: pOfficeColour = &H8000&
      Case 13: pOfficeColour = &H808000
      Case 14: pOfficeColour = &HFF0000
      Case 15: pOfficeColour = &H996666
      Case 16: pOfficeColour = &H808080
      
      Case 17: pOfficeColour = &HFF&
      Case 18: pOfficeColour = &H99FF&
      Case 19: pOfficeColour = &HCC99&
      Case 20: pOfficeColour = &H669933
      Case 21: pOfficeColour = &HCCCC33
      Case 22: pOfficeColour = &HFF6633
      Case 23: pOfficeColour = &H800080
      Case 24: pOfficeColour = &H999999
      
      Case 25: pOfficeColour = &HFF00FF
      Case 26: pOfficeColour = &HCCFF&
      Case 27: pOfficeColour = &HFFFF&
      Case 28: pOfficeColour = &HFF00&
      Case 29: pOfficeColour = &HFFFF00
      Case 30: pOfficeColour = &HFFCC00
      Case 31: pOfficeColour = &H663399
      Case 32: pOfficeColour = &HC0C0C0
      
      Case 33: pOfficeColour = &HCC99FF
      Case 34: pOfficeColour = &H99CCFF
      Case 35: pOfficeColour = &H99FFFF
      Case 36: pOfficeColour = &HCCFFCC
      Case 37: pOfficeColour = &HFFFFCC
      Case 38: pOfficeColour = &HFFCC99
      Case 39: pOfficeColour = &HFF99CC
      Case 40: pOfficeColour = &HFFFFFF
      End Select
   Else
      Select Case nIndex
      Case 1: pOfficeColour = &H0&
      Case 2: pOfficeColour = &H808080
      Case 3: pOfficeColour = &H80&
      Case 4: pOfficeColour = &H8080&
      Case 5: pOfficeColour = &H8000&
      Case 6: pOfficeColour = &H808000
      Case 7: pOfficeColour = &H800000
      Case 8: pOfficeColour = &H800080
      
      Case 9: pOfficeColour = &HFFFFFF
      Case 10: pOfficeColour = &HC0C0C0
      Case 11: pOfficeColour = &HFF&
      Case 12: pOfficeColour = &HFFFF&
      Case 13: pOfficeColour = &HFF00&
      Case 14: pOfficeColour = &HFFFF&
      Case 15: pOfficeColour = &HFF0000
      Case 16: pOfficeColour = &HFF0FF
      End Select
   End If
End Function
Private Function hexcolor(ByVal oColor As OLE_COLOR) As String
Dim sHex As String
   sHex = Hex$(oColor)
   If Len(sHex) < 6 Then
      hexcolor = "&H" & String$(6 - Len(sHex), "0") & sHex
   Else
      hexcolor = "&H" & sHex
   End If
End Function

Private Sub picMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
   X = X \ m_lS
   Y = Y \ m_lS
   Debug.Print X, Y
   If X >= 0 And X < 8 Then
      If Y >= 0 And Y < 5 Then
         i = Y * 8 + X + 1
         txtColour.Text = "#" & hexcolor(pOfficeColour(True, i))
      End If
   End If
End Sub

Private Sub picStandard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Long
   X = X \ m_lS
   Y = Y \ m_lS
   Debug.Print X, Y
   If X >= 0 And X < 8 Then
      If Y >= 0 And Y < 2 Then
         i = Y * 8 + X + 1
         txtColour.Text = "#" & hexcolor(pOfficeColour(False, i))
      End If
   End If
End Sub
