VERSION 5.00
Object = "*\AvbalDropDownForm6.vbp"
Begin VB.Form frmFindReplace 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   2115
   ClientLeft      =   3825
   ClientTop       =   3885
   ClientWidth     =   6390
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
   Icon            =   "frmFindReplace.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   Begin vbalDropDownForm6.vbalDropDownClient ddcDropDown 
      Align           =   1  'Align Top
      Height          =   135
      Left            =   0
      ToolTipText     =   "Drag to make this menu float"
      Top             =   0
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   238
      Caption         =   "Find and Replace"
   End
   Begin VB.PictureBox picControls 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1935
      ScaleWidth      =   6315
      TabIndex        =   1
      Top             =   180
      Width           =   6315
      Begin VB.ComboBox cboFindWhat 
         Height          =   315
         Left            =   1080
         TabIndex        =   10
         Top             =   120
         Width           =   3735
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "&Find Next"
         Default         =   -1  'True
         Height          =   375
         Left            =   4980
         TabIndex        =   9
         Top             =   60
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4980
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdReplace 
         Caption         =   "&Replace"
         Height          =   375
         Left            =   4980
         TabIndex        =   7
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton cmdReplaceAll 
         Caption         =   "Replace &All"
         Height          =   375
         Left            =   4980
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ComboBox cboReplaceWith 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Matc&h Case"
         Height          =   255
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Top             =   1020
         Width           =   2115
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Find &Whole Words Only"
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   3
         Top             =   1260
         Width           =   2115
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "Find in Se&lection Only"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   2
         Top             =   1500
         Width           =   2115
      End
      Begin VB.Label lblFindWhat 
         Caption         =   "Fi&nd what:"
         Height          =   255
         Left            =   60
         TabIndex        =   0
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label lblReplace 
         Caption         =   "Replace wi&th:"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   1275
      End
   End
End
Attribute VB_Name = "frmFindReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum EFRFindTypeOptions
   FR_DEFAULT = &H0
   FR_DOWN = &H1
   FR_WHOLEWORD = &H2
   FR_MATCHCASE = &H4&
End Enum

Public Event DoFind(ByVal sWhat As String, ByVal eOptions As EFRFindTypeOptions, ByVal bFindNext As Boolean, ByVal bSelection As Boolean)
Public Event DoReplace(ByVal sWhat As String, ByVal sWith As String, ByVal eOptions As EFRFindTypeOptions, ByVal bFindNext As Boolean, ByVal bSelection As Boolean, ByVal bReplaceAll As Boolean)

Private m_sFindHistory() As String
Private m_iFindHistoryCount As Long
Private m_sReplaceHistory() As String
Private m_iReplaceHistoryCount As Long
Private m_iMaxHistorySize As Long
Private m_sFindWhat As String
Private m_sReplaceWith As String

Private m_bMatchCase As Boolean
Private m_bWholeWord As Boolean
Private m_bSelection As Boolean

Public Enum EFRModeConstants
   efrFind = 0
   efrReplace = 1
End Enum
Private m_eMode As EFRModeConstants

Private m_bLoaded As Boolean

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

Public Property Let Mode(ByVal eMode As EFRModeConstants)
   m_eMode = eMode
   If (m_bLoaded) Then
      If (eMode = efrFind) Then
         cmdReplace.Caption = "&Replace..."
      Else
         cmdReplace.Caption = "&Replace"
      End If
      lblReplace.Visible = (eMode = efrReplace)
      cboReplaceWith.Visible = (eMode = efrReplace)
      cmdReplaceAll.Visible = (eMode = efrReplace)
   End If
End Property
Public Property Get Mode() As EFRModeConstants
   Mode = m_eMode
End Property

Public Property Let MaxHistorySize(ByVal lSize As Long)
   m_iMaxHistorySize = lSize
End Property
Public Property Get MaxHistorySize() As Long
   MaxHistorySize = m_iMaxHistorySize
End Property

Public Property Get FindHistoryCount() As Long
   FindHistoryCount = m_iFindHistoryCount
End Property
Public Property Get FindHistory(ByVal lIndex As Long) As String
   FindHistory = m_sFindHistory(lIndex)
End Property
Public Sub AddFindHistory(ByVal sText As String)
Dim i As Long
Dim sCurrent As String
   pAddHistory sText, m_sFindHistory(), m_iFindHistoryCount
   sCurrent = cboFindWhat.Text
   If (m_bLoaded) Then
      cboFindWhat.Clear
      For i = 1 To m_iFindHistoryCount
         cboFindWhat.AddItem m_sFindHistory(i)
      Next i
      cboFindWhat.Text = sCurrent
   End If
End Sub
Public Sub AddReplaceHistory(ByVal sText As String)
Dim i As Long
Dim sCurrent As String
   pAddHistory sText, m_sReplaceHistory(), m_iReplaceHistoryCount
   If (m_bLoaded) Then
      cboReplaceWith.Clear
      For i = 1 To m_iReplaceHistoryCount
         cboReplaceWith.AddItem m_sReplaceHistory(i)
      Next i
      cboReplaceWith.Text = sCurrent
   End If
End Sub
Public Property Let FindWhat(ByVal sText As String)
   m_sFindWhat = sText
   If (m_bLoaded) Then
      cboFindWhat.Text = sText
   End If
End Property
Public Property Let ReplaceWith(ByVal sText As String)
   
   m_sReplaceWith = sText
   If (m_bLoaded) Then
      cboReplaceWith.Text = sText
   End If
End Property
Private Sub pAddHistory( _
      ByVal sText As String, _
      ByRef sHistory() As String, _
      ByRef iCount As Long _
   )
Dim i As Long
Dim iFound As Long
Dim iMax As Long
   
   ' Check if already there:
   For i = 1 To iCount
      If (sHistory(i) = sText) Then
         iFound = i
         Exit For
      End If
   Next i
   
   ' Add this item as required:
   If (iFound) Then
      ' Swap iFound & 1:
      If (iFound <> 1) Then
         sHistory(iFound) = sHistory(1)
         sHistory(1) = sText
      End If
   Else
      ' Move all down and insert at 1:
      iMax = iCount + 1
      If (iMax > m_iMaxHistorySize) Then
         iMax = m_iMaxHistorySize
      End If
      If (iMax > iCount) Then
         iCount = iCount + 1
         ReDim Preserve sHistory(1 To iCount) As String
      End If
      For i = iMax To 2 Step -1
         sHistory(i) = sHistory(i - 1)
      Next i
      sHistory(1) = sText
   End If
   
End Sub


Public Property Get MatchCase() As Boolean
   MatchCase = m_bMatchCase
End Property
Public Property Let MatchCase(ByVal bState As Boolean)
   m_bMatchCase = bState
   If (m_bLoaded) Then
      chkOption(0).Value = Abs(bState)
   End If
End Property
Public Property Get WholeWord() As Boolean
   WholeWord = m_bWholeWord
End Property
Public Property Let WholeWord(ByVal bState As Boolean)
   m_bWholeWord = bState
   If (m_bLoaded) Then
      chkOption(1).Value = Abs(bState)
   End If
End Property
Public Property Get SelectionOnly() As Boolean
   SelectionOnly = m_bSelection
End Property
Public Property Let SelectionOnly(ByVal bState As Boolean)
   m_bSelection = bState
   If (m_bLoaded) Then
      chkOption(2).Value = Abs(bState)
   End If
End Property

Private Sub chkOption_Click(Index As Integer)
   Select Case Index
   Case 0
      m_bMatchCase = (chkOption(Index).Value = Checked)
   Case 1
      m_bWholeWord = (chkOption(Index).Value = Checked)
   Case 2
      m_bSelection = (chkOption(Index).Value = Checked)
   End Select
End Sub

Private Sub cmdCancel_Click()
   Me.Hide
End Sub

Private Sub cmdFindNext_Click()
Dim eOption As EFRFindTypeOptions
   eOption = peGetOption()
   RaiseEvent DoFind(cboFindWhat.Text, eOption, True, m_bSelection)
End Sub

Private Sub cmdReplace_Click()
Dim eOption As EFRFindTypeOptions
   If m_eMode = efrFind Then
      Mode = efrReplace
   Else
      eOption = peGetOption()
      RaiseEvent DoReplace(cboFindWhat.Text, cboReplaceWith.Text, eOption, True, m_bSelection, False)
   End If
End Sub

Private Sub cmdReplaceAll_Click()
Dim eOption As EFRFindTypeOptions
   eOption = peGetOption()
   RaiseEvent DoReplace(cboFindWhat.Text, cboReplaceWith.Text, eOption, True, m_bSelection, True)
End Sub

Private Function peGetOption() As EFRFindTypeOptions
Dim eOption As EFRFindTypeOptions
   eOption = FR_DOWN
   eOption = eOption Or (Abs(m_bMatchCase) * FR_MATCHCASE)
   eOption = eOption Or (Abs(m_bWholeWord) * FR_WHOLEWORD)
   peGetOption = eOption
End Function

Private Sub Form_Initialize()
   m_iMaxHistorySize = 5
End Sub

Private Sub Form_Load()
Dim i As Long

   ' Add History items:
   For i = 1 To m_iFindHistoryCount
      cboFindWhat.AddItem m_sFindHistory(i)
   Next i
   FindWhat = m_sFindWhat
   For i = 1 To m_iReplaceHistoryCount
      cboReplaceWith.AddItem m_sReplaceHistory(i)
   Next i
   ReplaceWith = m_sReplaceWith
   
   chkOption(0).Value = Abs(m_bMatchCase)
   chkOption(1).Value = Abs(m_bWholeWord)
   chkOption(2).Value = Abs(m_bSelection)
   
   m_bLoaded = True
   Mode = m_eMode
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (UnloadMode = vbFormControlMenu) Then
      ' Don't unload except in VB code.
      Cancel = True
      Me.Hide
   End If
End Sub

Private Sub ddcDropDown_AppActivate(ByVal bState As Boolean)
   ' Emulate Word - hide away floating
   ' toolwindows when we're not the focus
   ' app:
   If (bState) Then
      If ddcDropDown.ShowState = ewssFloating Then
         Me.Show
      End If
   Else
      If ddcDropDown.ShowState = ewssFloating Then
         Me.Hide
      End If
   End If
End Sub

Private Sub ddcDropDown_CaptionResize()
   ' Here you would resize your form/move the
   ' contents to accommodate the change in size
   ' of the caption:
   Me.Height = ddcDropDown.Height + picControls.Height + Me.Height - Me.ScaleHeight
End Sub

Private Sub ddcDropDown_CloseClick()
   ' User pressed the close button on the
   ' ToolWindow:
   Unload Me
End Sub


Private Sub Form_Resize()
   picControls.Move 0, ddcDropDown.Height
End Sub

