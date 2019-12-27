VERSION 5.00
Begin VB.UserControl SplitterPercent 
   Alignable       =   -1  'True
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   ControlContainer=   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8010
   ToolboxBitmap   =   "SplitterPercent.ctx":0000
   Begin VB.PictureBox SplitterBar 
      BorderStyle     =   0  'None
      Height          =   4536
      Left            =   3120
      ScaleHeight     =   4530
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   720
      Width           =   2652
   End
End
Attribute VB_Name = "SplitterPercent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'****************************************************************************************
' Orientation (Boolean, r/w) True: the splitter bar will be horizontal
' SplitPercent (Byte, r/w)   10-90 Percentage of the width of the control for first pane
' SplitterWidth (Byte, r/w) 0-*, width of the splitterbar
' SplitterColor (Long, r/w) Color value or constant that you want the splitterbar to be
' Child1 (Object, r/w) The control to act as pane1, the upper or left pane
' Child2 (Object, r/w) The control to act as pane2, the lower or right pane
' Be sure to set the Child1 and Child2 properties in the form load event to controls
' contained within the Splitterbar control.
'****************************************************************************************

'********************************
' Variables for properties
'********************************
Private mOrientation As SplitterOrientationEnum
Private mChild1        As Long
Private mChild2        As Long
Private mSplitPercent    As Single
Private mSplitterWidth      As Byte
Private mSplitterColor  As Long

Public Event Resize()

Public Property Get Orientation() As SplitterOrientationEnum
Attribute Orientation.VB_Description = "分割方向"
    Orientation = mOrientation
End Property

Public Property Let Orientation(val As SplitterOrientationEnum)
    mOrientation = val
    SplitterBar.MousePointer = IIf(val = SplitterOrientationHorizontal, vbSizeNS, vbSizeWE)
    PropertyChanged "Orientation"
    UserControl_Resize
End Property

Public Property Get Child1() As Object
Attribute Child1.VB_Description = "子控件1"
    Set Child1 = ObjectFromPtr(mChild1)
End Property

Public Property Set Child1(ctl As Object)
    mChild1 = ObjPtr(ctl)
    PropertyChanged "Child1"
'    UserControl_Resize
End Property

Public Property Get Child2() As Object
Attribute Child2.VB_Description = "子控件2"
    Set Child2 = ObjectFromPtr(mChild2)
End Property

Public Property Set Child2(ctl As Object)
    mChild2 = ObjPtr(ctl)
    PropertyChanged "Child2"
'    UserControl_Resize
End Property

Public Property Get SplitPercent() As Byte
Attribute SplitPercent.VB_Description = "分割比例"
    SplitPercent = mSplitPercent * 100
End Property

Public Property Let SplitPercent(val As Byte)
    mSplitPercent = val / 100
    PropertyChanged "SplitPercent"
    UserControl_Resize
End Property

Public Property Get SplitterWidth() As Byte
    SplitterWidth = mSplitterWidth
End Property

Public Property Let SplitterWidth(val As Byte)
    mSplitterWidth = val
    PropertyChanged "SplitterWidth"
    UserControl_Resize
End Property

Public Property Get SplitterColor() As Long
    SplitterColor = mSplitterColor
End Property

Public Property Let SplitterColor(val As Long)
    mSplitterColor = val
    SplitterBar.BackColor = val
    PropertyChanged "SplitterColor"
End Property

Private Sub UserControl_Initialize()
    mOrientation = SplitterOrientationHorizontal
    mSplitterColor = vbButtonFace
    mSplitterWidth = 50
    mSplitPercent = 50
    SplitterBar.BackColor = mSplitterColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Orientation = PropBag.ReadProperty("Orientation", SplitterOrientationHorizontal)
    SplitterWidth = PropBag.ReadProperty("SplitterWidth", 50)
    SplitterColor = PropBag.ReadProperty("SplitterColor", vbButtonFace)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Orientation", Orientation, SplitterOrientationHorizontal
    PropBag.WriteProperty "SplitterWidth", SplitterWidth, 50
    PropBag.WriteProperty "SplitterColor", SplitterColor, vbButtonFace
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
End Sub

Private Sub SplitterBar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With SplitterBar
        .BackColor = vbButtonText      ' Make the splitter visible
        .ZOrder
    End With
End Sub

Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        If mOrientation = SplitterOrientationHorizontal Then       ' horizontal figures
            y = SplitterBar.Top - (SplitterWidth - y)
            mSplitPercent = y / UserControl.Height
            SplitterBar.Move 0, y
        Else                                    ' vertical
            x = SplitterBar.Left - (SplitterWidth - x)
            mSplitPercent = x / UserControl.Width
            SplitterBar.Move x
        End If
        If mSplitPercent < 0.001 Then mSplitPercent = 0.001     ' Check if in range
        If mSplitPercent > 0.999 Then mSplitPercent = 0.999
'    UserControl_Resize                  ' update the panes
'        DoEvents
    End If
End Sub

Private Sub SplitterBar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SplitterBar.BackColor = SplitterColor  ' change the color back to normal
    UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

    Dim oChild1 As Object
    Dim oChild2 As Object
    If mChild1 > 0 Then
      Set oChild1 = Child1
    End If
    If mChild2 > 0 Then
      Set oChild2 = Child2
    End If

    Dim pane1 As Single
    Dim pane2 As Single
    Dim totwidth As Single
    Dim totheight As Single
    totwidth = UserControl.Width
    totheight = UserControl.Height
    If mOrientation = SplitterOrientationHorizontal Then
        pane1 = (totheight - SplitterWidth) * mSplitPercent
        pane2 = (totheight - SplitterWidth) * (1 - mSplitPercent)
        If Not oChild1 Is Nothing Then
            oChild1.Move 0, 0, totwidth, pane1
        End If
        If Not oChild2 Is Nothing Then
            oChild2.Move 0, pane1 + SplitterWidth, totwidth, pane2
        End If
        SplitterBar.Move 0, pane1, totwidth, SplitterWidth
    Else
        pane1 = (totwidth - SplitterWidth) * mSplitPercent
        pane2 = (totwidth - SplitterWidth) * (1 - mSplitPercent)
        If Not oChild1 Is Nothing Then
            oChild1.Move 0, 0, pane1, totheight
        End If
        If Not oChild2 Is Nothing Then
            oChild2.Move pane1 + SplitterWidth, 0, pane2, totheight
        End If
        SplitterBar.Move pane1, 0, SplitterWidth, totheight
    End If
    RaiseEvent Resize
End Sub
