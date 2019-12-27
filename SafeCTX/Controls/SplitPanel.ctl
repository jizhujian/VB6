VERSION 5.00
Begin VB.UserControl Splitter 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   ControlContainer=   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   1770
   ToolboxBitmap   =   "SplitPanel.ctx":0000
   Begin VB.PictureBox SplitterBar 
      Height          =   900
      Left            =   0
      ScaleHeight     =   840
      ScaleWidth      =   105
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
End
Attribute VB_Name = "Splitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'********************************
' Properties:
' HorizontalSplit (Boolean, r/w) True: the splitter
'               bar will be horizontal
' SplitPercent (Byte, r/w)   10-90 Percentage
'               of the width of the control for
'               first pane
' Control1 (Object, r/w) The control to act as
'               pane1, the upper or left pane
' Control2 (Object, r/w) The control to act as
'               pane2, the lower or right pane
' Be sure to set the Control1 and Control2
' properties in the form load event to controls
' contained within the SplitPanel control.
'
' The SplitPanel has a border during design time
' which disappears at run time.  Makes designing
' forms easier.
'********************************

'********************************
' Constants for properties
'********************************
'Private Const SplitWidth As Single = 80     ' width of splitterbar
Private Const nControl1  As String = "Control1"
Private Const nControl2  As String = "Control2"
'********************************
' Variables for properties
'********************************
Private mHorizontalSplit As Boolean
Private mControl1        As Object
Private mControl2        As Object
Private mSplitPercent    As Single
Private mSplitWidth      As Byte

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Sub AboutBox()
    About
End Sub

'********************************
' Read-Write Properties
'********************************
Public Property Get HorizontalSplit() As Boolean
    HorizontalSplit = mHorizontalSplit
End Property
Public Property Let HorizontalSplit(val As Boolean)
    mHorizontalSplit = val
    SplitterBar.MousePointer = IIf(HorizontalSplit, vbSizeNS, vbSizeWE)
    PropertyChanged "HorizontalSplit"
    UserControl_Resize
End Property

Public Property Get Control1() As Object
    Set Control1 = mControl1
End Property
Public Property Set Control1(ctl As Object)
    Set mControl1 = ctl
    PropertyChanged nControl1
    UserControl_Resize
End Property

Public Property Get Control2() As Object
    Set Control2 = mControl2
End Property
Public Property Set Control2(ctl As Object)
    Set mControl2 = ctl
    PropertyChanged nControl2
    UserControl_Resize
End Property

Public Property Get SplitPercent() As Byte
    SplitPercent = mSplitPercent * 100
End Property
Public Property Let SplitPercent(val As Byte)
    mSplitPercent = val / 100
    PropertyChanged "SplitPercent"
    UserControl_Resize
End Property

Public Property Get SplitWidth() As Byte
    SplitWidth = mSplitWidth
End Property
Public Property Let SplitWidth(val As Byte)
    mSplitWidth = val
    PropertyChanged "SplitWidth"
End Property

Private Sub UserControl_Initialize()
    SplitterBar.BorderStyle = vbBSNone
End Sub

'********************************
' Set up the defaults
'********************************
Private Sub UserControl_InitProperties()
    HorizontalSplit = False
    SplitPercent = 50
'    SplitWidth = 80
'    MousePointer = vbSizeNWSE
End Sub

'********************************
' Reload design-time settings
'********************************
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    HorizontalSplit = PropBag.ReadProperty("HorizontalSplit", False)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
    SplitWidth = PropBag.ReadProperty("SplitWidth", 80)
End Sub

'********************************
' Save design-time settings
'********************************
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HorizontalSplit", HorizontalSplit, False
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
    PropBag.WriteProperty "SplitWidth", SplitWidth, 80
End Sub

'********************************
' These next three subs handle the actual
' dragging of the splitterbar.  The panes
' are updated when the mouse button is
' released.
'********************************
Private Sub SplitterBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With SplitterBar
        .BackColor = &H8000000C     ' Make the splitter visible
        .ZOrder
    End With
End Sub

Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mHorizontalSplit Then        ' horizontal figures
            Y = SplitterBar.Top - (SplitWidth - Y)
            mSplitPercent = Y / UserControl.Height
            SplitterBar.Move 0, Y
        Else                                    ' vertical
            X = SplitterBar.Left - (SplitWidth - X)
            mSplitPercent = X / UserControl.Width
            SplitterBar.Move X
        End If
        If mSplitPercent < 0.001 Then mSplitPercent = 0.001     ' Check if in range
        If mSplitPercent > 0.999 Then mSplitPercent = 0.999
    End If
End Sub

Private Sub SplitterBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'SplitterBar.ZOrder 99
    SplitterBar.BackColor = &H8000000F  ' change the color back to normal
    UserControl_Resize                  ' update the panes
End Sub

'********************************
' The resize event is where it get's ugly
' Here we must figure out the sizes and
' positions of everything based on the splitter
' position, and the controls properties, then
' set everything
'********************************
Private Sub UserControl_Resize()
    On Error Resume Next
    
    If UserControl.Ambient.UserMode Then    ' get rid of border in run mode
        UserControl.BorderStyle = vbBSNone
    End If
    
    Dim pane1 As Single
    Dim pane2 As Single
    Dim totwidth As Single
    Dim totheight As Single
    totwidth = UserControl.Width
    totheight = UserControl.Height
    If mHorizontalSplit Then
        pane1 = (totheight - SplitWidth) * mSplitPercent
        pane2 = (totheight - SplitWidth) * (1 - mSplitPercent)
        mControl1.Move 0, 0, totwidth, pane1
        mControl2.Move 0, pane1 + SplitWidth, totwidth, pane2
        SplitterBar.Move 0, pane1, totwidth, SplitWidth
    Else
        pane1 = (totwidth - SplitWidth) * mSplitPercent
        pane2 = (totwidth - SplitWidth) * (1 - mSplitPercent)
        mControl1.Move 0, 0, pane1, totheight
        mControl2.Move pane1 + SplitWidth, 0, pane2, totheight
        SplitterBar.Move pane1, 0, SplitWidth, totheight
    End If
    mControl1.Refresh
    mControl2.Refresh
End Sub

