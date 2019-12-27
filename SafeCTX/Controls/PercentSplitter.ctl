VERSION 5.00
Begin VB.UserControl PercentSplitter 
   Alignable       =   -1  'True
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   ControlContainer=   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8010
   ToolboxBitmap   =   "PercentSplitter.ctx":0000
   Begin VB.PictureBox SplitterBar 
      Height          =   4656
      Left            =   3480
      ScaleHeight     =   4590
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   600
      Width           =   132
   End
End
Attribute VB_Name = "PercentSplitter"
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
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'********************************
' Variables for properties
'********************************
Private mOrientation As Boolean
Private mChild1        As Long
Private mChild2        As Long
Private mSplitPercent    As Single
Private mSplitterWidth      As Byte
Private mSplitterColor  As Long

Public Event Resize()

Public Enum splitterOrientation
    splitVertical = False
    splitHorizontal = True
End Enum

Public Property Get Orientation() As splitterOrientation
    Orientation = mOrientation
End Property

Public Property Let Orientation(val As splitterOrientation)
    mOrientation = val
    SplitterBar.MousePointer = IIf(Orientation, vbSizeNS, vbSizeWE)
    PropertyChanged "Orientation"
    UserControl_Resize
End Property

Public Property Get Child1() As Object
    Set Child1 = ObjectFromPtr(mChild1)
End Property

Public Property Set Child1(ctl As Object)
    mChild1 = PtrFromObject(ctl)
    PropertyChanged "Child1"
    UserControl_Resize
End Property

Public Property Get Child2() As Object
    Set Child2 = ObjectFromPtr(mChild2)
End Property

Public Property Set Child2(ctl As Object)
    mChild2 = PtrFromObject(ctl)
    PropertyChanged "Child2"
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
    Orientation = False
    SplitPercent = 50
    SplitterWidth = 50
    mSplitterColor = vbButtonFace
    With SplitterBar
        .BorderStyle = vbBSNone
        .BackColor = mSplitterColor
        .Width = SplitterWidth
    End With
End Sub

'Private Sub UserControl_Paint()
'    If SplitterColor <> vbButtonFace Then
'        If SplitterColor = 0 Then
'            SplitterBar.BackColor = vbButtonFace
'        Else
'            SplitterBar.BackColor = SplitterColor
'        End If
'    Else
'        SplitterBar.BackColor = vbButtonFace
'    End If
'End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Orientation = PropBag.ReadProperty("Orientation", False)
    SplitPercent = PropBag.ReadProperty("SplitPercent", 50)
    SplitterWidth = PropBag.ReadProperty("SplitterWidth", 50)
    SplitterColor = PropBag.ReadProperty("SplitterColor", vbButtonFace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Orientation", Orientation, False
    PropBag.WriteProperty "SplitPercent", SplitPercent, 50
    PropBag.WriteProperty "SplitterWidth", SplitterWidth, 50
    PropBag.WriteProperty "SplitterColor", SplitterColor, vbButtonFace
End Sub

Private Sub SplitterBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With SplitterBar
        .BackColor = vbButtonText      ' Make the splitter visible
        .ZOrder
    End With
End Sub

Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If mOrientation Then        ' horizontal figures
            Y = SplitterBar.Top - (SplitterWidth - Y)
            mSplitPercent = Y / UserControl.Height
            SplitterBar.Move 0, Y
        Else                                    ' vertical
            X = SplitterBar.Left - (SplitterWidth - X)
            mSplitPercent = X / UserControl.Width
            SplitterBar.Move X
        End If
        If mSplitPercent < 0.001 Then mSplitPercent = 0.001     ' Check if in range
        If mSplitPercent > 0.999 Then mSplitPercent = 0.999
    UserControl_Resize                  ' update the panes
        DoEvents
    End If
End Sub

Private Sub SplitterBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SplitterBar.BackColor = SplitterColor  ' change the color back to normal
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next

'    If UserControl.Ambient.UserMode Then    ' get rid of border in run mode
'        UserControl.BorderStyle = vbBSNone
'    End If

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
    If mOrientation Then
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
'    If Not oChild1 Is Nothing Then
'        oChild1.Refr1esh
'    End If
'    If Not oChild2 Is Nothing Then
'        oChild2.Refresh
'    End If
    RaiseEvent Resize
End Sub

Private Function ObjectFromPtr(ByVal lPtr As Long) As Object
    Dim oThis As Object

    ' Turn the pointer into an illegal, uncounted interface
    CopyMemory oThis, lPtr, 4
    ' Do NOT hit the End button here! You will crash!
    ' Assign to legal reference
    Set ObjectFromPtr = oThis
    ' Still do NOT hit the End button here! You will still crash!
    ' Destroy the illegal reference
    CopyMemory oThis, 0&, 4
    ' OK, hit the End button if you must--you'll probably still crash,
    ' but this will be your code rather than the uncounted reference!
End Function

Private Function PtrFromObject(ByRef oThis) As Long
    ' Return the pointer to this object:
    PtrFromObject = ObjPtr(oThis)
End Function

