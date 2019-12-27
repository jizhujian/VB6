VERSION 5.00
Begin VB.UserControl vbSplitter 
   Alignable       =   -1  'True
   ClientHeight    =   2145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   150
   ControlContainer=   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   150
   ToolboxBitmap   =   "vbSplitter.ctx":0000
   Begin VB.PictureBox SplitterBar 
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   0
      ScaleHeight     =   2130
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   0
      Width           =   120
   End
End
Attribute VB_Name = "vbSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'***********************************************************************************************
'Splitter Control for VB
'Copyright: ©2001 Matthew Hood, Dragon Wery Development
'Author(s): Matthew Hood Email: DragonWeyrDev@Yahoo.com
'Description: This control creates a splitter bar with 2 resizable panels.
' The panels are adjustable by a specifying the size of the Child1 panel
' through the SetPanelSize method. It also includes minimum and maxmimu size parameters.
'Credits: Thanks to Mark Joyal for his great SplitterControl on which this is based.
' His control provides a way to resize based by size percentage.
' You can download his control from
' http://www.planetsourcecode.com/xq/ASP/txtCodeId.5855/lngWId.1/qx/vb/scripts/ShowCode.htm
'***********************************************************************************************
'Revision History:
'[Matthew Hood]
'   01/18/01 - New
'***********************************************************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

'***********************************************************************************************
'Types/Enumerations
'***********************************************************************************************
Public Enum SplitOrientationConstants
    vbSplitHorizontal = 1
    vbSplitVertical = 2
End Enum

Public Enum SplitBorderStyleConstants
    vbStyleNone = 0
    vbStyleFixedSingle = 1
End Enum
'***********************************************************************************************
'API Declarations
'***********************************************************************************************
'***********************************************************************************************
'Private Variables/Constants
'***********************************************************************************************
Private mMaxSize As Single 'Maximum size of Child1.
Private mMinSize As Single 'Minimum size of Child1.
Private mAutoResize As Boolean 'Allows panels to be resized on the fly.
Private mOrientation As Integer 'Splitter orientation.
Private mPanels(1) As Long 'Panel objects.
Private mPanelSize As Single 'Size of Child1.
Private mProportional As Boolean 'Adust the properties proportionately if orientation is change.
Private mSplitterColor As Long 'Splitter color.
Private mSplitterWidth As Single 'Splitter size.
Private mSelectedColor As Long 'Selected splitter color.
'***********************************************************************************************
'Public Events
'***********************************************************************************************
'This Resize event allows other controls to respond to resizing the panels.
Public Event Resize()
'***********************************************************************************************
'Public Properties/Constants
'***********************************************************************************************
'The BorderStyle property specifies the border style of the control.
Public Property Get BorderStyle() As SplitBorderStyleConstants
    BorderStyle = UserControl.BorderStyle
End Property
Public Property Let BorderStyle(ByVal Value As SplitBorderStyleConstants)
On Error Resume Next
    'Save the property.
    UserControl.BorderStyle = Value
    PropertyChanged "BorderStyle"
End Property

'The Child1 property specifies the 1st panel object.
Public Property Get Child1() As Object
On Error Resume Next
    Set Child1 = ObjectFromPtr(mPanels(0))
End Property
Public Property Set Child1(ByRef Obj As Object)
On Error Resume Next
    'Save the property.
    mPanels(0) = PtrFromObject(Obj)
    PropertyChanged "Child1"
End Property

'The Child2 property specifies the 2nd panel object.
Public Property Get Child2() As Object
On Error Resume Next
    Set Child2 = ObjectFromPtr(mPanels(1))
End Property
Public Property Set Child2(ByRef Obj As Object)
    'Save the property.
    mPanels(1) = PtrFromObject(Obj)
    PropertyChanged "Child2"
End Property

'The MaxSize property specifies the maximum size the Child1 panel.
Public Property Get MaxSize() As Single
    MaxSize = mMaxSize
End Property
Public Property Let MaxSize(ByVal Value As Single)
On Error Resume Next
    'Set to 0 to have no maxiumum size.

    'Make sure the Value parameter is a valid value.
    If Value < 0 Then Value = 0
    Select Case mOrientation
        Case vbSplitHorizontal
            If Value > UserControl.ScaleHeight Then Value = UserControl.ScaleHeight
        Case vbSplitVertical
            If Value > UserControl.ScaleWidth Then Value = UserControl.ScaleWidth
    End Select

    'Make sure the MaxSize is not less than the MinSize.
    If Value <> 0 And Value < mMinSize Then Value = mMinSize

    'Save the property.
    mMaxSize = Value
    PropertyChanged "MaxSize"

    'Resize the panels if the MaxSize is less than the current Child1 panel size.
    If Value < mPanelSize And Value <> 0 Then PanelSize = Value
End Property

'The MinSize property specifies the minimum size the Child1 panel.
Public Property Get MinSize() As Single
    MinSize = mMinSize
End Property
Public Property Let MinSize(ByVal Value As Single)
On Error Resume Next
    'Set to 0 to have no mimiumum size.

    'Make sure the Value parameter is a valid value.
    If Value < 0 Then Value = 0
    Select Case mOrientation
        Case vbSplitHorizontal
            If Value > UserControl.ScaleHeight Then Value = UserControl.ScaleHeight
        Case vbSplitVertical
            If Value > UserControl.ScaleWidth Then Value = UserControl.ScaleWidth
    End Select

    'Make sure the MinSize is not greater than the MaxSize.
    If Value <> 0 And Value > mMaxSize And mMaxSize <> 0 Then Value = mMaxSize

    'Save the property.
    mMinSize = Value
    PropertyChanged "MinSize"
    
    'Resize the panels if the MinSize is greater than the current Child1 panel size.
    If Value > PanelSize Then PanelSize = Value
End Property

'The AutoResize property specifies wether or not to resize the panels
'during or after the splitter is moved.
Public Property Get AutoResize() As Boolean
    AutoResize = mAutoResize
End Property
Public Property Let AutoResize(ByVal Value As Boolean)
On Error Resume Next
    'Save the property.
    mAutoResize = Value
    PropertyChanged "AutoResize"
End Property

'The Orientation property specifies the splitter orientation.
Public Property Get Orientation() As SplitOrientationConstants
    Orientation = mOrientation
End Property
Public Property Let Orientation(ByVal Value As SplitOrientationConstants)
On Error Resume Next
    'Make sure Value parameter is a valid value.
    If Value <> 1 And Value <> 2 Then Value = vbSplitVertical

    'Change to the appropriate sizer pointer and reset the panel size to 1/2 the control size.
    Select Case Value
        Case vbSplitHorizontal
            SplitterBar.MousePointer = vbSizeNS
'            mPanelSize = UserControl.ScaleHeight / 2
        Case vbSplitVertical
            SplitterBar.MousePointer = vbSizeWE
'            mPanelSize = UserControl.ScaleWidth / 2
    End Select

    'Save the property.
    mOrientation = Value
    PropertyChanged "Orientation"
    
    'Resize the panels.
    PanelSize = mPanelSize
End Property

'Get's the Child1 panel's size.
Public Property Get PanelSize() As Single
    PanelSize = mPanelSize
End Property
Public Property Let PanelSize(ByVal Value As Single)
On Error Resume Next
    'Make sure the Value parameter is a valid value.
    If Value < 0 Then Value = 0
    
    'Make sure Value parameter is not greater than the total control size.
    Select Case mOrientation
        Case vbSplitHorizontal
            If Value > UserControl.ScaleHeight - mSplitterWidth Then
                Value = UserControl.ScaleHeight - mSplitterWidth
            End If
        Case vbSplitVertical
            If Value > UserControl.ScaleWidth - mSplitterWidth Then
                Value = UserControl.ScaleWidth - mSplitterWidth
            End If
    End Select

    'Save the property.
    mPanelSize = Value
    PropertyChanged "PanelSize"

    'Resize the panels.
    Call UserControl_Resize
End Property

'The SelectedColor property specifies the color of the splitter bar when it is selected.
Public Property Get SelectedColor() As SystemColorConstants
    SelectedColor = mSelectedColor
End Property
Public Property Let SelectedColor(ByVal Value As SystemColorConstants)
On Error Resume Next
    'Save the property.
    mSelectedColor = Value
    PropertyChanged "SelectedColor"
End Property

'The SplitterColor property specifies the color of the splitter bar.
Public Property Get SplitterColor() As SystemColorConstants
    SplitterColor = mSplitterColor
End Property
Public Property Let SplitterColor(ByVal Value As SystemColorConstants)
On Error Resume Next
    'Change the splitter color.
    SplitterBar.BackColor = Value

    'Save the property.
    mSplitterColor = Value
    PropertyChanged "SplitterColor"
End Property

'The SplitterWidth property specifies the width of the splitter bar.
Public Property Get SplitterWidth() As Single
    SplitterWidth = mSplitterWidth
End Property
Public Property Let SplitterWidth(ByVal Value As Single)
    'Make sure the Value parameter is a valid value.
    If Value < 0 Then Value = 0
'    Select Case mOrientation
'        Case vbSplitHorizontal
'            If SplitterBar.Top + Value > UserControl.ScaleHeight Then
'                mPanelSize = UserControl.ScaleHeight - Value
'            End If
'        Case vbSplitVertical
'            If SplitterBar.Left + Value > UserControl.ScaleWidth Then
'                mPanelSize = UserControl.ScaleWidth - Value
'            End If
'    End Select

    'Save the property.
    mSplitterWidth = Value
    PropertyChanged "SplitterWidth"
    
    'Resize the panels to adust for the the new width.
    PanelSize = mPanelSize
End Property
'***********************************************************************************************
'Public Methods
'***********************************************************************************************
'The ForceResize method forces the control to resize the panels.
Public Sub ForceResize()
On Error Resume Next
    'Force a resize of the panels.
    PanelSize = mPanelSize
End Sub
'***********************************************************************************************
'Private Methods
'***********************************************************************************************
'***********************************************************************************************
'Load/Unload Events
'***********************************************************************************************
Private Sub UserControl_Initialize()
On Error Resume Next
    'Set property default values.
    mAutoResize = True
    mOrientation = vbSplitHorizontal
    mSelectedColor = vbButtonText
    mSplitterColor = vbButtonFace
    With SplitterBar
        .BackColor = mSplitterColor
        .Width = 50
    End With
    mSplitterWidth = SplitterBar.Width
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
    'Read the property values.
    BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    MaxSize = PropBag.ReadProperty("MaxSize", 0)
    MinSize = PropBag.ReadProperty("MinSize", 0)
    AutoResize = PropBag.ReadProperty("AutoResize", True)
    Orientation = PropBag.ReadProperty("Orientation", vbSplitHorizontal)
    SelectedColor = PropBag.ReadProperty("SelectedColor", vbButtonText)
    SplitterColor = PropBag.ReadProperty("SplitterColor", vbButtonFace)
    SplitterWidth = PropBag.ReadProperty("SplitterWidth", 50)
    PanelSize = PropBag.ReadProperty("PanelSize", 0)
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
    'Save the property values.
    PropBag.WriteProperty "BorderStyle", BorderStyle, 0
    PropBag.WriteProperty "MaxSize", MaxSize, 0
    PropBag.WriteProperty "MinSize", MinSize, 0
    PropBag.WriteProperty "AutoResize", AutoResize, True
    PropBag.WriteProperty "Orientation", Orientation, vbSplitHorizontal
    PropBag.WriteProperty "SelectedColor", SelectedColor, vbButtonText
    PropBag.WriteProperty "SplitterColor", SplitterColor, vbButtonFace
    PropBag.WriteProperty "SplitterWidth", SplitterWidth, 50
    PropBag.WriteProperty "PanelSize", PanelSize, 0
End Sub
'***********************************************************************************************
'Resize Events
'***********************************************************************************************
Private Sub UserControl_Resize()
On Error Resume Next
    Dim sngLeft As Single 'Child2 panel left value.
    Dim sngTop As Single 'Child2 panel top value.
    Dim sngSize As Single 'Child2 panel size.
    Dim sngWidth As Single 'Control's scalewidth value.
    Dim sngHeight As Single 'Control's scaleheight value.

    Dim oChild1 As Object
    Dim oChild2 As Object
    If mPanels(0) > 0 Then
      Set oChild1 = Child1
    End If
    If mPanels(1) > 0 Then
      Set oChild2 = Child2
    End If

    'Get the control size.
    sngWidth = UserControl.ScaleWidth
    sngHeight = UserControl.ScaleHeight

    'Resize the panels.
    Select Case mOrientation
        Case vbSplitHorizontal
            SplitterBar.Move 0, mPanelSize, sngWidth, mSplitterWidth

            'Resize the Child1 panel.
            If Not oChild1 Is Nothing Then
                oChild1.Move 0, 0, sngWidth, mPanelSize
            End If

            'Set the Child2 panel location & size.
            sngTop = mPanelSize + mSplitterWidth
            sngSize = sngHeight - (mPanelSize + mSplitterWidth)

            'Resize the Child2 panel.
            If Not oChild2 Is Nothing Then
                oChild2.Move 0, sngTop, sngWidth, sngSize
            End If
        Case vbSplitVertical
            SplitterBar.Move mPanelSize, 0, mSplitterWidth, sngHeight

            'Resize the Child1 panel.
            If Not oChild1 Is Nothing Then
                oChild1.Move 0, 0, mPanelSize, sngHeight
            End If

            'Set the Child2 panel location & size.
            sngLeft = mPanelSize + mSplitterWidth
            sngSize = sngWidth - (mPanelSize + mSplitterWidth)

            'Resize the Child2 panel.
            If Not oChild2 Is Nothing Then
                oChild2.Move sngLeft, 0, sngSize, sngHeight
            End If
    End Select

    'Refresh the panels.
'    If Not oChild1 Is Nothing Then
'        oChild1.Refresh
'    End If
'    If Not oChild2 Is Nothing Then
'        oChild2.Refresh
'    End If
    
    'Raise the Resize event.
    RaiseEvent Resize
End Sub
'***********************************************************************************************
'Focus Events
'***********************************************************************************************
'***********************************************************************************************
'Click Events
'***********************************************************************************************
'***********************************************************************************************
'Keyboard/Mouse Events
'***********************************************************************************************
Private Sub SplitterBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    'Set the splitter color to the selected to the selected state and bring it to the top.
    With SplitterBar
        .BackColor = SelectedColor
        .ZOrder
    End With
End Sub
Private Sub SplitterBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Dim sngMin As Single 'Mimium size. (adjusted)
    Dim sngMax As Single 'Maximum size. (adjusted)
    Dim sngPos As Single 'New position value for splitter.

    'Exit if the left mouse button is not pressed.
    If Button <> vbLeftButton Then Exit Sub

    Select Case mOrientation
        Case vbSplitHorizontal
            'Define the new splitter position.
            sngPos = SplitterBar.Top + Y
            
            'Get the Child1 min/max sizes.
            sngMin = mMinSize
            sngMax = mMaxSize
            If sngMax = 0 Then sngMax = UserControl.ScaleHeight - mSplitterWidth

            'Make sure splitter is positioned inside the control.
            If sngPos < sngMin Then
                sngPos = sngMin
            ElseIf sngPos > sngMax Then
                sngPos = sngMax
            End If

            'Move the splitter.
            SplitterBar.Move 0, sngPos

            'Resize panels if AutoResize is enabled.
            If mAutoResize Then PanelSize = sngPos
        Case vbSplitVertical
            'Define the new splitter position.
            sngPos = SplitterBar.Left + X
            
            'Get the Child1 min/max sizes.
            sngMin = mMinSize
            sngMax = mMaxSize
            If sngMax = 0 Then sngMax = UserControl.ScaleWidth - mSplitterWidth

            'Make sure splitter is positioned inside the control.
            If sngPos < sngMin Then
                sngPos = sngMin
            ElseIf sngPos > sngMax Then
                sngPos = sngMax
            End If

            'Move the splitter.
            SplitterBar.Move sngPos
            
            'Resize panels if AutoResize is enabled.
            If mAutoResize Then PanelSize = sngPos
    End Select

    'Refresh the panels.
End Sub
Private Sub SplitterBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    'Set the splitter color back to the unselected state.
    SplitterBar.BackColor = mSplitterColor
    
    'Exit if On-The-Fly resizing is enabled. (Panels already sized.)
    If AutoResize Then Exit Sub

    'Resize the panels.
    Select Case mOrientation
        Case vbSplitHorizontal
            PanelSize = SplitterBar.Top
        Case vbSplitVertical
            PanelSize = SplitterBar.Left
    End Select
End Sub
'***********************************************************************************************
'Change/Validation Events
'***********************************************************************************************
'***********************************************************************************************
'Control Events
'***********************************************************************************************
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
