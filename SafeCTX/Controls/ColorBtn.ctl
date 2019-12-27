VERSION 5.00
Begin VB.UserControl ColorButton 
   ClientHeight    =   1830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1635
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   1635
   ToolboxBitmap   =   "ColorBtn.ctx":0000
   Begin VB.PictureBox picPopup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   45
      ScaleHeight     =   1815
      ScaleWidth      =   1170
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1170
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'API Types
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
     lpCustColorControls As Long
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

'Custom Type
Private Type ColorControl
    Color As Long
    Area As RECT
    Interior As RECT
End Type


'API Functions
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long

'API Constants
Private Const DT_CENTER = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_LEFT = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKEN = &HA
Private Const BDR_RAISED = &H5
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_TOOLWINDOW = &H80
Private Const CC_FULLOPEN = &H2
Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const CC_RGBINIT = &H1

'Private Members
Private ColorControls(21) As ColorControl
Private nCurrentControl As Integer
Private mctlCancel As Object

'Event Declarations:
Event Click() 'MappingInfo=picPopup,picPopup,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."

'Default Property Values:
Const m_def_SelectedColor = vbWhite

'Property Variables:
Dim m_SelectedColor As OLE_COLOR

Public Sub AboutBox()
    About
End Sub

Private Sub picPopup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim na As Integer
    
    'Find out whether we're over a well or button
    For na = 1 To 21
        If PtInRect(ColorControls(na).Area, X, Y) Then
            'Only need to do anything if we're over a new one
            If na = nCurrentControl Then
                Exit Sub
            Else
                If nCurrentControl > 0 And nCurrentControl <> 21 Then
                    DrawEdge picPopup.hdc, ColorControls(nCurrentControl).Area, BDR_SUNKEN, BF_RECT
                    picPopup.Line (ColorControls(nCurrentControl).Interior.Left, ColorControls(nCurrentControl).Interior.Top)-(ColorControls(nCurrentControl).Interior.Right - 1, ColorControls(nCurrentControl).Interior.Bottom - 1), ColorControls(nCurrentControl).Color, BF
                End If
                If na < 21 Then 'it's not the 'Other' button
                    ForeColor = vbWhite
                    picPopup.Line (ColorControls(na).Area.Left, ColorControls(na).Area.Top)-(ColorControls(na).Area.Right - 1, ColorControls(na).Area.Bottom - 1), vbBlack, B
                    picPopup.Line (ColorControls(na).Area.Left + 1, ColorControls(na).Area.Top + 1)-(ColorControls(na).Area.Right - 2, ColorControls(na).Area.Bottom - 2), vbWhite, B
                    picPopup.Line (ColorControls(na).Area.Left + 2, ColorControls(na).Area.Top + 2)-(ColorControls(na).Area.Right - 3, ColorControls(na).Area.Bottom - 3), vbBlack, B
                End If
                'remember the current control
                nCurrentControl = na
                Exit For
            End If
        End If
    Next

End Sub

Private Sub picPopup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim rct As RECT

    GetClientRect picPopup.hWnd, rct
    
    'check to see that we have a valid Current Control
    If PtInRect(rct, X / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY) And nCurrentControl > 0 And Button = vbLeftButton Then
        'user has selected a color so set m_SelectedColor
        m_SelectedColor = ColorControls(nCurrentControl).Color
        'User feedback that color has changed
        RaiseEvent Click
    End If
    HidePopUp

End Sub

Private Sub picPopup_Paint()
    
    Dim rct As RECT
    Dim na As Integer
    
    'Draw the 3D edge
    GetClientRect picPopup.hWnd, rct
    DrawEdge picPopup.hdc, rct, BDR_RAISED, BF_RECT
    
    'Draw the color wells and button
    For na = 1 To 20 'the wells
        DrawEdge picPopup.hdc, ColorControls(na).Area, BDR_SUNKEN, BF_RECT
        picPopup.Line (ColorControls(na).Interior.Left, ColorControls(na).Interior.Top)-(ColorControls(na).Interior.Right - 1, ColorControls(na).Interior.Bottom - 1), ColorControls(na).Color, BF
    Next na
    'The button
    DrawEdge picPopup.hdc, ColorControls(21).Area, BDR_RAISED, BF_RECT
    'print the button caption
    DrawText picPopup.hdc, "Other...", Len("Other..."), ColorControls(21).Interior, DT_CENTER

End Sub

Private Sub UserControl_Click()
    If Not picPopup.Visible Then
        ShowPopUp
    Else
        'a second click is treated as escape
        HidePopUp
    End If
End Sub

Private Sub UserControl_EnterFocus()
    'repaint so we get a focus rectangle
    UserControl_Paint
End Sub

Private Sub UserControl_ExitFocus()
    'Although in most circumstances the popup window will have already been
    'hidden before this, we check here just in case.
    If picPopup.Visible Then HidePopUp
    UserControl_Paint 'hide the focus rectangle
End Sub

Private Sub UserControl_Initialize()

    'Set the parent and window style for the popup picturebox
    'set style to Toolwindow so after we've set parent to the Desktop
    'the popup doesn't show in the Taskbar
    SetWindowLong picPopup.hWnd, GWL_EXSTYLE, WS_EX_TOOLWINDOW
    SetParent picPopup.hWnd, 0
    
    'for simplicity, set the scalemode of the popup picturebox to pixels
    picPopup.ScaleMode = vbPixels
    'Set up the co-ordinates and colors for the Color Controls on the popup
    InitColorControls

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the popup to show/hide
    'Since this is a button, we show the popup if the user presses
    'Space and hide it if the user presses escape.
    If KeyCode = vbKeySpace And (Shift = 0) Then
        UserControl_Click
    ElseIf KeyCode = vbKeyEscape And picPopup.Visible Then
        HidePopUp
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DrawButton
End Sub

Private Sub UserControl_Paint()
    DrawButton
End Sub

Private Sub HidePopUp()

    'This procedure is called whenever the popup window needs to be hidden.
    If GetCapture = picPopup.hWnd Then
        ReleaseCapture
    End If
    picPopup.Visible = False
    DoEvents
    UserControl_Paint
    'restore he Cancel property if we have a default 'Cancel' control
    If Not mctlCancel Is Nothing Then
        mctlCancel.Cancel = True
    End If
End Sub

Private Sub ShowPopUp()
    
    'This procedure is called whenever the popup needs to be shown.
    
    Dim ileft As Long
    Dim iTop As Long
    Dim ctlRect As RECT
    Dim na As Long
    
    'Determine position for pop up window
    'We want to show the popup below the control, but if we can't we'll show it above
    GetWindowRect UserControl.hWnd, ctlRect 'screen rectange of the control
    If ctlRect.Bottom + (picPopup.Height / Screen.TwipsPerPixelX) > Screen.Height / Screen.TwipsPerPixelY Then
        'put it above
        iTop = (ctlRect.Top - (picPopup.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
    Else
        'put it below
        iTop = ctlRect.Bottom * Screen.TwipsPerPixelY
    End If
    'If the popup window is as wide as, or wider than the control, we want to align
    'it to the left edge of the control.  Otherwise, we align it to the right.  If
    'we're too far to the right, we push it back left.
    If (ctlRect.Right - ctlRect.Left) > picPopup.Width / Screen.TwipsPerPixelX Then
        'try to align to the right of the control
        If ctlRect.Right > Screen.Width / Screen.TwipsPerPixelX Then
            ileft = Screen.Width - picPopup.Width
        Else
            ileft = ctlRect.Right * Screen.TwipsPerPixelX - picPopup.Width
        End If
        'Check we haven't gone outside the left edge of the screen
        If ileft < 0 Then ileft = 0
    Else
        'try to align to the left
        If ctlRect.Left < 0 Then
            ileft = 0
        Else
            ileft = ctlRect.Left * Screen.TwipsPerPixelX
        End If
        'Check we haven't gone outside the left edge of the screen
        If ileft + picPopup.Width > Screen.Width Then ileft = Screen.Width - picPopup.Width
    End If
    
    With picPopup
        .Top = iTop
        .Left = ileft
        .Visible = True
        .ZOrder
    End With
    picPopup_Paint
    DoEvents
    UserControl_Paint
    'Capture the mouse so we get all subsequent mouse clicks
    SetCapture picPopup.hWnd
    'store the 'Cancel' control so we stop Escape from firing the default 'Cancel' button
    On Error Resume Next
    For na = 0 To UserControl.ParentControls.Count - 1
        If UserControl.ParentControls(na).Cancel Then
            If Err = False Then
                Set mctlCancel = UserControl.ParentControls(na)
                mctlCancel.Cancel = False
                Exit For
            End If
            Err = False
        End If
    Next na
    On Error GoTo 0
End Sub

Private Sub picPopUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lColor As Long
    
    'see if a box has been selected...
    If nCurrentControl > 0 Then
        If nCurrentControl = 21 Then
            'user has selected the 'Other' button - show the color dialog
            HidePopUp
            If GetColor(lColor) Then
                m_SelectedColor = lColor
                UserControl_Paint
                RaiseEvent Click
            End If
        End If
    End If

End Sub

Private Sub InitColorControls()

    'Initialize the array of pseudo controls for our popup
    Dim rct As RECT
    Dim lX As Long
    Dim lY As Long
    Dim lBorderOffset As Long
    Dim lCurControl As Long
    Dim lWellsWidth As Long
    Dim lWellsHeight As Long

    lBorderOffset = 4
    lCurControl = 1
    
    'create the boxes and load them into the array
    For lY = 0 To 4
        For lX = 0 To 3
            'create the control's rectangle
            With rct
                .Left = lBorderOffset + lX * 22
                .Top = lBorderOffset + lY * 22
                .Right = lBorderOffset + lX * 22 + 20
                .Bottom = lBorderOffset + lY * 22 + 20
            End With
            lWellsWidth = lBorderOffset + lX * 22 + 20
            lWellsHeight = lBorderOffset + lY * 22 + 20
            ColorControls(lCurControl).Area = rct
            'create the Interior where we paint the color/print he caption
            With rct
                .Left = lBorderOffset + lX * 22 + 2
                .Top = lBorderOffset + lY * 22 + 2
                .Right = lBorderOffset + lX * 22 + 18
                .Bottom = lBorderOffset + lY * 22 + 18
            End With
            ColorControls(lCurControl).Interior = rct
            lCurControl = lCurControl + 1
        Next
    Next
    
    'the 'Other' button
    With rct
        .Left = 4
        .Top = lWellsHeight + lBorderOffset
        .Right = lBorderOffset + 2 * 22 + 20
        .Bottom = lWellsHeight + lBorderOffset + 21
    End With
    ColorControls(21).Area = rct
    
    With rct
        .Left = 8
        .Top = lWellsHeight + lBorderOffset + 4
        .Right = lBorderOffset + 2 * 22 + 16
        .Bottom = lWellsHeight + lBorderOffset + 18
    End With
    ColorControls(21).Interior = rct
    
    'set the height and width of the popup to the area used
    With picPopup
        .Height = (lBorderOffset + lWellsHeight + lBorderOffset + 21) * Screen.TwipsPerPixelY
        .Width = (lWellsWidth + lBorderOffset) * Screen.TwipsPerPixelX
    End With
    
    'set the color of the wells
    ColorControls(1).Color = RGB(255, 255, 255)
    ColorControls(2).Color = RGB(0, 0, 0)
    ColorControls(3).Color = RGB(192, 192, 192)
    ColorControls(4).Color = RGB(128, 128, 128)
    ColorControls(5).Color = RGB(255, 0, 0)
    ColorControls(6).Color = RGB(128, 0, 0)
    ColorControls(7).Color = RGB(255, 255, 0)
    ColorControls(8).Color = RGB(128, 128, 0)
    ColorControls(9).Color = RGB(0, 255, 0)
    ColorControls(10).Color = RGB(0, 128, 0)
    ColorControls(11).Color = RGB(0, 255, 255)
    ColorControls(12).Color = RGB(0, 128, 128)
    ColorControls(13).Color = RGB(0, 0, 255)
    ColorControls(14).Color = RGB(0, 0, 128)
    ColorControls(15).Color = RGB(255, 0, 255)
    ColorControls(16).Color = RGB(128, 0, 128)
    ColorControls(17).Color = RGB(192, 220, 192)
    ColorControls(18).Color = RGB(166, 202, 240)
    ColorControls(19).Color = RGB(255, 251, 240)
    ColorControls(20).Color = RGB(160, 160, 164)

End Sub

Private Function GetColor(lColor As Long) As Long

    Dim hMem As Long
    Dim rtn As Long
    Dim na As Long
    Dim cc As ChooseColor
    Dim alCustColorControls(15) As Long
    Dim lCustColorSize As Long
    Dim lCustColorAddress As Long
    
    'Set up the ChhoseColor structure
    cc.lStructSize = Len(cc)
    cc.hWndOwner = hWnd
    cc.rgbResult = m_SelectedColor
    
    'Initialize the custom colors to white
    For na = 0 To UBound(alCustColorControls)
        alCustColorControls(na) = &HFFFFFF
    Next na
    lCustColorSize = Len(alCustColorControls(0)) * 16
    hMem = GlobalAlloc(GHND, lCustColorSize)
    lCustColorAddress = GlobalLock(hMem)
    
    ' Copy custom colors to  global memory
    Call CopyMemory(ByVal lCustColorAddress, alCustColorControls(0), ByVal lCustColorSize)
    cc.lpCustColorControls = lCustColorAddress
    
    'set the flags - full open
    cc.flags = CC_FULLOPEN Or CC_RGBINIT
    
    'call the ChooseColor API function
    If ChooseColor(cc) = 1 Then
        lColor = cc.rgbResult 'return the selected color
        GetColor = True
    Else
        GetColor = False
    End If
    
    'clean up
    rtn = GlobalUnlock(hMem)
    hMem = GlobalFree(hMem)

End Function

Private Sub DrawButton(Optional bPressed As Boolean = False)
    
    Dim rct As RECT
    
    UserControl.Cls
    
    'Paint the interior with the selected color
    UserControl.Line ((5 - bPressed) * Screen.TwipsPerPixelX, (4 - bPressed) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - ((16 + bPressed) * Screen.TwipsPerPixelX), UserControl.ScaleHeight - ((5 + bPressed) * Screen.TwipsPerPixelY)), vbBlack, B
    UserControl.Line ((6 - bPressed) * Screen.TwipsPerPixelX, (5 - bPressed) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - ((17 + bPressed) * Screen.TwipsPerPixelX), UserControl.ScaleHeight - ((6 + bPressed) * Screen.TwipsPerPixelY)), m_SelectedColor, BF

    'add the separator
    UserControl.Line (UserControl.ScaleWidth - ((12 + bPressed)) * Screen.TwipsPerPixelX, (4 - bPressed) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - ((12 + bPressed) * Screen.TwipsPerPixelX), UserControl.ScaleHeight - ((5 + bPressed) * Screen.TwipsPerPixelY)), vbButtonShadow, BF
    UserControl.Line (UserControl.ScaleWidth - ((11 + bPressed)) * Screen.TwipsPerPixelX, (4 - bPressed) * Screen.TwipsPerPixelY)-(UserControl.ScaleWidth - ((11 + bPressed) * Screen.TwipsPerPixelX), UserControl.ScaleHeight - ((5 + bPressed) * Screen.TwipsPerPixelY)), vb3DHighlight, BF
    
    'Add the combo symbol
    UserControl.Line (UserControl.ScaleWidth - ((9 + bPressed)) * Screen.TwipsPerPixelX, Int(UserControl.ScaleHeight / 2) - ((1 + bPressed) * Screen.TwipsPerPixelY))-((UserControl.ScaleWidth - ((5 + bPressed)) * Screen.TwipsPerPixelX), (Int(UserControl.ScaleHeight / 2)) - ((1 + bPressed) * Screen.TwipsPerPixelY)), vbBlack, BF
    UserControl.Line (UserControl.ScaleWidth - ((8 + bPressed)) * Screen.TwipsPerPixelX, Int(UserControl.ScaleHeight / 2) - (bPressed * Screen.TwipsPerPixelY))-((UserControl.ScaleWidth - ((6 + bPressed)) * Screen.TwipsPerPixelX), (Int(UserControl.ScaleHeight / 2)) - (bPressed * Screen.TwipsPerPixelY)), vbBlack, BF
    UserControl.Line (UserControl.ScaleWidth - ((7 + bPressed)) * Screen.TwipsPerPixelX, Int(UserControl.ScaleHeight / 2) + ((1 - bPressed) * Screen.TwipsPerPixelY))-((UserControl.ScaleWidth - ((7 + bPressed)) * Screen.TwipsPerPixelX), (Int(UserControl.ScaleHeight / 2)) + ((1 - bPressed) * Screen.TwipsPerPixelY)), vbBlack, BF
    
    'Draw the border
    GetClientRect UserControl.hWnd, rct
    If bPressed Then
        DrawEdge UserControl.hdc, rct, BDR_SUNKEN, BF_RECT
    Else
        DrawEdge UserControl.hdc, rct, BDR_RAISED, BF_RECT
    End If
    
    'Draw a focus rectangle if we need one
    If Ambient.UserMode And GetFocus = UserControl.hWnd And Not picPopup.Visible Then
        GetClientRect UserControl.hWnd, rct
        With rct
            .Left = .Left + 3
            .Right = .Right - 3
            .Top = .Top + 3
            .Bottom = .Bottom - 3
        End With
        DrawFocusRect UserControl.hdc, rct
    End If

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get SelectedColor() As OLE_COLOR
    SelectedColor = m_SelectedColor
End Property

Public Property Let SelectedColor(ByVal New_SelectedColor As OLE_COLOR)
    m_SelectedColor = New_SelectedColor
    PropertyChanged "SelectedColor"
    UserControl_Paint
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_SelectedColor = m_def_SelectedColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_SelectedColor = PropBag.ReadProperty("SelectedColor", m_def_SelectedColor)
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("SelectedColor", m_SelectedColor, m_def_SelectedColor)
End Sub

