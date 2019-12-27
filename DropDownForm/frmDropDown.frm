VERSION 5.00
Object = "{4A3A29A4-F2E3-11D3-B06C-00500427A693}#4.0#0"; "vbalDDFm6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form frmDropDown 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1935
   ClientLeft      =   4695
   ClientTop       =   4050
   ClientWidth     =   4365
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
   ScaleHeight     =   1935
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvwDropDown 
      Height          =   1515
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2672
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   3720
      Top             =   1260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDropDown.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDropDown.frx":27B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbalDropDownForm6.vbalDropDownClient ddcDropDown 
      Align           =   1  'Align Top
      Height          =   90
      Left            =   0
      ToolTipText     =   "Drag to make this menu float"
      Top             =   0
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   159
      AllowTearOff    =   0   'False
      AllowResize     =   0   'False
   End
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced..."
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const elviReportFullRowSelect = &H20            '// applies to report mode only

Private Const LVM_FIRST = &H1000                   '// ListView messages
Private Const LVM_GETITEMRECT = (LVM_FIRST + 14)
Private Const LVM_GETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 55)
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE = (LVM_FIRST + 54) '// optional wParam == mask
Private Declare Function SendMessageByLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private m_sValue As String
Private m_bCancel As Boolean
Private m_iNameCount As Long
Private m_sFirstName() As String
Private m_sSurName() As String

Public Event Change()
Public Event CloseUp()

Public Property Let ShowState(ByVal eState As EWindowShowState)
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ddcDropDOwn.ShowState = eState
End Property
Public Property Get ShowState() As EWindowShowState
   ' This is to allow the parent form
   ' to control the current drop-down state:
   ShowState = ddcDropDOwn.ShowState
End Property

Private Sub cmdAdvanced_Click()
   MsgBox "Show Advanced Dialog Here"
End Sub

Private Sub ddcDropDown_AppActivate(ByVal bState As Boolean)
   ' Emulate Word - hide away floating
   ' toolwindows when we're not the focus
   ' app:
   If (bState) Then
      If ddcDropDOwn.ShowState = ewssFloating Then
         Me.Show
      End If
   Else
      If ddcDropDOwn.ShowState = ewssFloating Then
         Me.Hide
      End If
   End If
End Sub

Private Sub ddcDropDown_CaptionResize()
   ' Here you would resize your form/move the
   ' contents to accommodate the change in size
   ' of the caption:
   
End Sub

Private Sub ddcDropDown_CloseClick()
   ' User pressed the close button on the
   ' ToolWindow:
   Unload Me
End Sub

Public Property Get Cancelled() As Boolean
   Cancelled = m_bCancel
End Property
Public Property Get Value() As String
   Value = m_sValue
End Property
Public Property Let Value(ByVal sValue As String)
Dim itmX As ListItem
   m_sValue = sValue
   If ddcDropDOwn.ShowState = ewssDropped Then
      For Each itmX In lvwDropDown.ListItems
         With itmX
            If sValue = .Text & " " & .SubItems(1) & " " & .SubItems(2) Then
               itmX.Selected = True
               Exit For
            End If
         End With
      Next
   End If
End Property

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

Private Property Get vowel(Optional ByVal iHowMany As Long = 1) As String
Dim i As Long
Dim sOut As String
Dim s As String
   For i = 1 To iHowMany
      Do
         s = Chr$(Asc("a") + Rnd * 25)
      Loop While Not (isin(s, "a", "e", "i", "o", "u"))
      sOut = sOut & s
   Next i
   vowel = sOut
End Property

Private Property Get consonant(Optional ByVal iHowMany As Long = 1) As String
Dim i As Long
Dim sOut As String
Dim s As String
   For i = 1 To iHowMany
      Do
         s = Chr$(Asc("a") + Rnd * 25)
      Loop While isin(s, "a", "e", "i", "o", "u")
      sOut = sOut & s
   Next i
   consonant = sOut
End Property
Private Function isin(ByVal s As String, ParamArray vOptions() As Variant) As Boolean
Dim i As Long
   For i = LBound(vOptions) To UBound(vOptions)
      If (s = vOptions(i)) Then
         isin = True
         Exit Function
      End If
   Next i
End Function

'Private Sub Form_Activate()
'   lvwDropDown.SetFocus
'End Sub

Private Sub Form_Load()
   m_bCancel = True
   
   ' Set up the ListView so it has a multi-select style:
   Dim i As Long, itmX As ListItem
   Dim lStyle As Long
   lvwDropDown.SmallIcons = ilsIcons
   lvwDropDown.View = lvwReport
   lvwDropDown.ColumnHeaders.Add , , "Title", 16 * Screen.TwipsPerPixelX
   lvwDropDown.ColumnHeaders.Add , , "First Name", 64 * Screen.TwipsPerPixelX
   lvwDropDown.ColumnHeaders.Add , , "Surname", 64 * Screen.TwipsPerPixelX
   GenerateNames
   For i = 1 To m_iNameCount
      Set itmX = lvwDropDown.ListItems.Add(, , "Mr", , 1)
      itmX.SubItems(1) = m_sFirstName(i)
      itmX.SubItems(2) = m_sSurName(i)
   Next i
   lStyle = SendMessageByLong(lvwDropDown.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
   lStyle = lStyle Or elviReportFullRowSelect
   SendMessageByLong lvwDropDown.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0, lStyle
   
End Sub
Private Sub GenerateNames()
Dim i As Long
Static bFirst As Boolean
   If Not bFirst Then
      m_iNameCount = Rnd * 20 + 20
      ReDim m_sFirstName(1 To m_iNameCount) As String
      ReDim m_sSurName(1 To m_iNameCount) As String
      For i = 1 To m_iNameCount
         m_sFirstName(i) = UCase$(consonant) & vowel & consonant(2) & vowel & consonant
         m_sSurName(i) = UCase$(consonant) & vowel(2) & consonant(2) & vowel & consonant(2)
      Next i
      bFirst = True
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   RaiseEvent CloseUp
End Sub

Private Sub Form_Resize()
   lvwDropDown.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - cmdAdvanced.Height - 4 * Screen.TwipsPerPixelY
   cmdAdvanced.Move cmdAdvanced.Left, Me.ScaleHeight - cmdAdvanced.Height - 2 * Screen.TwipsPerPixelY
End Sub

Private Sub lvwDropDown_Click()
   With lvwDropDown.SelectedItem
      m_sValue = .Text & " " & .SubItems(1) & " " & .SubItems(2)
      RaiseEvent Change
      m_bCancel = False
   End With
   Unload Me
End Sub


Private Sub lvwDropDown_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim colX As ColumnHeader
   For Each colX In lvwDropDown.ColumnHeaders
      If Not ColumnHeader Is colX Then
         colX.Tag = ""
      End If
   Next
   If ColumnHeader.Tag = "ASC" Then
      ColumnHeader.Tag = "DESC"
   Else
      ColumnHeader.Tag = "ASC"
   End If
   lvwDropDown.SortKey = ColumnHeader.Index - 1
   If (ColumnHeader.Tag = "ASC") Then
      lvwDropDown.SortOrder = lvwAscending
   Else
      lvwDropDown.SortOrder = lvwDescending
   End If
   lvwDropDown.Sorted = True
End Sub

Private Sub lvwDropDown_ItemClick(ByVal Item As MSComctlLib.ListItem)
   With lvwDropDown.SelectedItem
      m_sValue = .Text & " " & .SubItems(1) & " " & .SubItems(2)
      RaiseEvent Change
   End With
End Sub

Private Sub lvwDropDown_KeyUp(KeyCode As Integer, Shift As Integer)
   If (KeyCode = vbKeyReturn) Then
      lvwDropDown_Click
   ElseIf (KeyCode = vbKeyEscape) Then
      Unload Me
   End If
End Sub

Private Sub lvwDropDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'SetCapture lvwDropDown.hwnd
End Sub

Private Sub lvwDropDown_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tR As RECT
Dim tP As POINTAPI
Dim i As Long
Dim lhWNd As Long
   Debug.Print X, Y
   lhWNd = lvwDropDown.hwnd
   GetCursorPos tP
   ScreenToClient lvwDropDown.hwnd, tP
   For i = 1 To lvwDropDown.ListItems.Count
      SendMessage lhWNd, LVM_GETITEMRECT, i - 1, tR
      If tR.Top > tP.Y Then
         Exit For
      Else
         If tR.Left <= tP.X And tR.Right >= tP.X Then
            If tR.Top <= tP.Y And tR.Bottom >= tP.Y Then
               lvwDropDown.ListItems(i).Selected = True
               Exit For
            End If
         End If
      End If
   Next i
End Sub

Private Sub lvwDropDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   lvwDropDown_ItemClick lvwDropDown.SelectedItem
End Sub
