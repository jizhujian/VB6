VERSION 5.00
Begin VB.UserControl vbalDropDownClient 
   Alignable       =   -1  'True
   BackColor       =   &H80000002&
   CanGetFocus     =   0   'False
   ClientHeight    =   3315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3315
   ScaleWidth      =   4800
   ToolboxBitmap   =   "vbalDropDownClient.ctx":0000
End
Attribute VB_Name = "vbalDropDownClient"
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
Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Enum EHitTestAreas
   HTERROR = (-2)
   HTTRANSPARENT = (-1)
   HTNOWHERE = 0
   HTCLIENT = 1
   HTCAPTION = 2
   HTSYSMENU = 3
   HTGROWBOX = 4
   HTMENU = 5
   HTHSCROLL = 6
   HTVSCROLL = 7
   HTMINBUTTON = 8
   HTMAXBUTTON = 9
   HTLEFT = 10
   HTRIGHT = 11
   HTTOP = 12
   HTTOPLEFT = 13
   HTBOTTOM = 15
   HTBOTTOMLEFT = 16
   HTBOTTOMRIGHT = 17
   HTBORDER = 18
End Enum

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_MOVE = &HF010&
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10

Private Const WM_ACTIVATE = &H6
Private Const WM_DESTROY = &H2
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_NCHITTEST = &H84&
Private Const WM_NCLBUTTONDOWN = &HA1

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function DrawFrameControl Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Private Const DFC_CAPTION = 1

Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONHELP = &H4
Private Const DFCS_CAPTIONCLOSE = &H0

Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const DFCS_CHECKED = &H400
'#if(WINVER >= =&H0500)
Private Const DFCS_TRANSPARENT = &H800
Private Const DFCS_HOT = &H1000
'#endif /* WINVER >= =&H0500 */
Private Const DFCS_ADJUSTRECT = &H2000
Private Const DFCS_FLAT = &H4000
Private Const DFCS_MONO = &H8000

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_LEFT = &H0
Private Const DT_NOPREFIX = &H800
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_TITLEBAR = DT_NOPREFIX Or DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS Or DT_MODIFYSTRING

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT = 1

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long

'/* flags for DrawCaption */
Private Enum DCFlags
   DC_ACTIVE = &H1&
   DC_SMALLCAP = &H2&
   DC_ICON = &H4&
   DC_TEXT = &H8&
   DC_INBUTTON = &H10&
'#if(WINVER >= 0x0500)
   DC_GRADIENT = &H20&
'#endif /* WINVER >= 0x0500 */
End Enum
Private Declare Function DrawCaption Lib "user32" (ByVal hwnd As Long, ByVal lhDC As Long, tR As RECT, ByVal eFlag As DCFlags) As Long

Implements ISubclass

Public Enum EWindowShowState
   ewssHidden
   ewssDropped
   ewssFloating
End Enum

Private m_bAllowTearOff As Boolean
Private m_bNonClientMove As Boolean
Private m_eShowState As EWindowShowState
Private m_iHysteresis  As Long
Private m_lTearOffHeight As Long
Private m_sCaption As String
Private m_bAllowResize As Boolean

Private m_bSubclassing As Boolean
Private m_hWnd As Long
Private m_bRunTime As Boolean
Private m_tPMouseDown As POINTAPI
Private m_bMouseDown As Boolean
Private m_bMouseOver As Boolean
Private m_hDC As Long, m_hBmp As Long, m_hBmpOld As Long
Private m_lWidthDC As Long, m_lHeightDC As Long
Private m_bAppActive As Boolean
Private m_tEnterSize As RECT
Private m_bSizing As Boolean

Private WithEvents m_cMT As cMouseTrack
Attribute m_cMT.VB_VarHelpID = -1
Private m_cNCM As cNCMetrics
Private WithEvents m_cSM As cSizeMoveHelper
Attribute m_cSM.VB_VarHelpID = -1

Private m_bButton As Boolean
Private m_bButtonDown As Boolean
Private m_bButtonOver As Boolean
Private m_tButtonR As RECT

Public Event AppActivate(ByVal bState As Boolean)
Public Event DeactivateForm()
Public Event CaptionResize()
Public Event Moving(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Public Event RightClick(ByVal x As Single, ByVal y As Single)
Public Event Sizing(ByRef lLeft As Long, ByRef lTop As Long, ByRef lWidth As Long, ByRef lHeight As Long)

Public Property Get Sizing() As Boolean
   Sizing = m_bSizing
End Property

Public Property Get hwnd() As Long
   hwnd = UserControl.hwnd
End Property

Private Sub DrawTitleBar()
Dim tR As RECT
Dim lWidth As Long
Dim lHeight As Long
Dim lhDC As Long
Dim bDrawDirect As Boolean
Dim hBr As Long
Dim lBarColor As Long
Dim lCapColor As Long
Dim hFont As Long
Dim hFontOld As Long
Dim lStyle As Long

   If m_eShowState > ewssHidden Then
      GetClientRect UserControl.hwnd, tR
      lWidth = Abs(tR.Right - tR.Left)
      lHeight = Abs(tR.Bottom - tR.Top)
         
      ' Memory DC for draw speed:
      If lWidth > m_lWidthDC Or lHeight > m_lHeightDC Then
         pRebuildDC lWidth, lHeight
      End If
      lhDC = m_hDC
      If lhDC = 0 Then lhDC = UserControl.hDC: bDrawDirect = True
      
      ' Draw gradient if possible
      lStyle = DC_SMALLCAP Or DC_TEXT Or DC_GRADIENT
      If m_eShowState = ewssDropped Then
         If m_bMouseOver Then
            ' active tbar
            lStyle = lStyle Or DC_ACTIVE
            lBarColor = (vbActiveTitleBar And &H1F&)
            lCapColor = (vbTitleBarText And &H1F&)
         Else
            lBarColor = (vbInactiveTitleBar And &H1F&)
            lCapColor = (vbInactiveCaptionText And &H1F&)
         End If
         
      ElseIf m_eShowState = ewssFloating Then
         ' draw active if app is active
         If m_bAppActive Then
            lStyle = lStyle Or DC_ACTIVE
            lBarColor = (vbActiveTitleBar And &H1F&)
            lCapColor = (vbTitleBarText And &H1F&)
         Else
            lBarColor = (vbInactiveTitleBar And &H1F&)
            lCapColor = (vbInactiveCaptionText And &H1F&)
         End If
      End If
        '
      If m_eShowState = ewssFloating Then
         DrawCaption m_hWnd, m_hDC, tR, lStyle
         pDrawCloseButton m_hDC
      Else
         hBr = GetSysColorBrush(lBarColor)
         FillRect lhDC, tR, hBr
         DeleteObject hBr
      End If
      
      
      If Not bDrawDirect Then
         BitBlt UserControl.hDC, 0, 0, lWidth, lHeight, m_hDC, 0, 0, vbSrcCopy
      End If
   End If
   
End Sub
Private Sub pDrawCloseButton(ByVal lhDC As Long)
Dim lH As Long
Dim lStyle As Long
Dim lType As Long
Dim bEnabled As Boolean
Dim tR As RECT

   If m_bButton Then
      bEnabled = True
      m_cNCM.GetMetrics
      lH = m_cNCM.CaptionHeight - 2
      GetClientRect UserControl.hwnd, tR
      tR.Left = tR.Right - lH - 2
      tR.Top = tR.Top + 2
      tR.Right = tR.Left + lH
      tR.Bottom = tR.Top + lH - 2
      LSet m_tButtonR = tR
      lType = DFC_CAPTION
      lStyle = DFCS_CAPTIONCLOSE
      If (m_bButtonDown And m_bButtonOver) Then
         lStyle = lStyle Or DFCS_PUSHED
      End If
      If Not (bEnabled) Then
         lStyle = lStyle Or DFCS_INACTIVE
      End If
      DrawFrameControl lhDC, tR, lType, lStyle
   End If
   
End Sub

Private Sub pRebuildDC(ByVal lWidth As Long, ByVal lHeight As Long)
   If lWidth > m_lWidthDC Then
      m_lWidthDC = lWidth
   End If
   If lHeight > m_lHeightDC Then
      m_lHeightDC = lHeight
   End If
   If Not m_hBmpOld = 0 Then
      SelectObject m_hDC, m_hBmpOld
      m_hBmpOld = 0
   End If
   If Not m_hBmp = 0 Then
      DeleteObject m_hBmp
      m_hBmp = 0
   End If
   If m_hDC = 0 Then
      m_hDC = CreateCompatibleDC(UserControl.hDC)
   End If
   If Not m_hDC = 0 Then
      m_hBmp = CreateCompatibleBitmap(UserControl.hDC, m_lWidthDC, m_lHeightDC)
      If Not m_hBmp = 0 Then
         m_hBmpOld = SelectObject(m_hDC, m_hBmp)
      End If
   End If
   If m_hBmpOld = 0 Then
      If Not m_hBmp = 0 Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
      If Not m_hDC = 0 Then
         DeleteDC m_hDC
         m_hDC = 0
      End If
      m_lWidthDC = 0
      m_lHeightDC = 0
   End If
   
End Sub
Public Property Get ShowState() As EWindowShowState
   ShowState = m_eShowState
End Property
Public Property Let ShowState(ByVal eState As EWindowShowState)
   If Not m_eShowState = eState Then
      m_eShowState = eState
      m_bMouseOver = False
      m_bMouseDown = False
      Select Case m_eShowState
      Case ewssHidden
         m_bButton = False
         PostMessage m_hWnd, WM_CLOSE, 0, 0
      Case ewssFloating
         m_bAppActive = True
         m_cNCM.GetMetrics
         m_bButton = True
         UserControl.Extender.Height = m_cNCM.CaptionHeight * Screen.TwipsPerPixelY
         RaiseEvent CaptionResize
         DrawTitleBar
      Case ewssDropped
         m_bAppActive = True
         m_bButton = False
         UserControl.Extender.Height = m_lTearOffHeight * Screen.TwipsPerPixelY
         RaiseEvent CaptionResize
         DrawTitleBar
      End Select
   End If
End Property

Public Property Get AllowResize() As Boolean
   AllowResize = m_bAllowResize
End Property
Public Property Let AllowResize(ByVal bState As Boolean)
   m_bAllowResize = bState
   PropertyChanged "AllowResize"
End Property

Public Property Get Caption() As String
   Caption = m_sCaption
End Property
Public Property Let Caption(ByVal sCaption As String)
   m_sCaption = sCaption
   If m_bRunTime Then
      SetWindowText ParenthWnd, sCaption
   End If
   PropertyChanged "Caption"
End Property

Public Property Get AllowTearOff() As Boolean
   AllowTearOff = m_bAllowTearOff
End Property
Public Property Let AllowTearOff(ByVal bState As Boolean)
   m_bAllowTearOff = bState
   If m_bRunTime Then
      If Not bState Then
         UserControl.Extender.Height = 0
      End If
      RaiseEvent CaptionResize
   End If
   PropertyChanged "AllowTearOff"
End Property

Public Property Get MoveOnFormMouseDown() As Boolean
   MoveOnFormMouseDown = m_bNonClientMove
End Property
Public Property Let MoveOnFormMouseDown(ByVal bState As Boolean)
   m_bNonClientMove = bState
   PropertyChanged "MoveOnFormMouseDown"
End Property

Public Property Get ParenthWnd() As Long
Dim lParenthWnd As Long
   lParenthWnd = UserControl.Parent.hwnd
   If Not (lParenthWnd = m_hWnd) Then
      pTerminate
      pInitialise
   End If
   ParenthWnd = lParenthWnd
End Property

Private Sub pInitialise()
   
   pTerminate
   
   m_hWnd = UserControl.Parent.hwnd
   m_bRunTime = UserControl.Ambient.UserMode

   If m_bRunTime Then
      AttachMessage Me, m_hWnd, WM_NCHITTEST
      AttachMessage Me, m_hWnd, WM_ACTIVATE
      AttachMessage Me, m_hWnd, WM_ACTIVATEAPP
      AttachMessage Me, m_hWnd, WM_DESTROY
      m_bSubclassing = True
      
      Set m_cMT = New cMouseTrack
      m_cMT.AttachMouseTracking Me
            
      Set m_cSM = New cSizeMoveHelper
      m_cSM.Attach m_hWnd
      
   End If
   
End Sub
Private Sub pTerminate()

   If m_bSubclassing Then
      'Debug.Print "Terminate"
      DetachMessage Me, m_hWnd, WM_NCHITTEST
      DetachMessage Me, m_hWnd, WM_ACTIVATE
      DetachMessage Me, m_hWnd, WM_ACTIVATEAPP
      DetachMessage Me, m_hWnd, WM_DESTROY
      
      ' Clear up DC and bitmap:
      If Not m_hBmpOld = 0 Then
         SelectObject m_hDC, m_hBmpOld
         m_hBmpOld = 0
      End If
      If Not m_hBmp = 0 Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
      If Not m_hDC = 0 Then
         DeleteDC m_hDC
         m_hDC = 0
      End If
      
      ' Stop mouse tracking:
      m_cMT.DetachMouseTracking
      Set m_cMT = Nothing
            
      ' Stop size/move checking:
      m_cSM.Detach
      Set m_cSM = Nothing
   
      m_bSubclassing = False
   End If
   
End Sub

Private Property Let ISubClass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
   '
End Property

Private Property Get ISubClass_MsgResponse() As SSubTimer6.EMsgResponse
   Select Case CurrentMessage
   Case WM_ACTIVATE
      ISubClass_MsgResponse = emrPostProcess
   Case Else
      ISubClass_MsgResponse = emrPreprocess
   End Select
End Property

Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tR As RECT
Dim tP As POINTAPI

   Select Case iMsg
   Case WM_NCHITTEST
      GetClientRect ParenthWnd, tR
      GetCursorPos tP
      ScreenToClient ParenthWnd, tP
      If m_bNonClientMove Then
         If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
            ISubClass_WindowProc = HTCAPTION
            Exit Function
         End If
      End If
      If m_eShowState = ewssDropped Or Not (m_bAllowResize) Then
         ISubClass_WindowProc = HTCLIENT
      Else
         ' Do the default:
         ISubClass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
      
   Case WM_ACTIVATE
      If wParam = 0 Then
         RaiseEvent DeactivateForm
         If m_eShowState = ewssDropped Then
            ShowState = ewssHidden
         End If
      Else
         ' active:
         Debug.Print "Activate"
      End If
      
   Case WM_ACTIVATEAPP
      If (wParam = 0) Then
         ' deactivate
         m_bAppActive = False
         DrawTitleBar
         RaiseEvent AppActivate(False)
      Else
         ' activate
         m_bAppActive = True
         DrawTitleBar
         RaiseEvent AppActivate(True)
      End If
   
   Case WM_DESTROY
      'Debug.Print "WM_DESTROY"
      pTerminate
      
   End Select
   
End Function

Private Sub m_cMT_MouseHover(Button As MouseButtonConstants, Shift As ShiftConstants, x As Single, y As Single)
   m_cMT.StartMouseTracking
End Sub

Private Sub m_cMT_MouseLeave()
   m_bMouseOver = False
   DrawTitleBar
End Sub

Private Sub m_cSM_EnterSizeMove()
   m_bSizing = True
   GetWindowRect m_hWnd, m_tEnterSize
End Sub

Private Sub m_cSM_ExitSizeMove()
   m_bSizing = False
End Sub

Private Sub m_cSM_Moving(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
   If m_eShowState = ewssDropped Then
      lLeft = m_tEnterSize.Left
      lTop = m_tEnterSize.Top
      lWidth = m_tEnterSize.Right - m_tEnterSize.Left
      lHeight = m_tEnterSize.Bottom - m_tEnterSize.Top
   ElseIf m_eShowState = ewssFloating Then
      RaiseEvent Moving(lLeft, lTop, lWidth, lHeight)
   End If
End Sub

Private Sub m_cSM_Sizing(lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long)
   If m_eShowState = ewssDropped Then
      lLeft = m_tEnterSize.Left
      lTop = m_tEnterSize.Top
      lWidth = m_tEnterSize.Right - m_tEnterSize.Left
      lHeight = m_tEnterSize.Bottom - m_tEnterSize.Top
   ElseIf m_eShowState = ewssFloating Then
      If m_bAllowResize Then
         RaiseEvent Sizing(lLeft, lTop, lWidth, lHeight)
      Else
         lLeft = m_tEnterSize.Left
         lTop = m_tEnterSize.Top
         lWidth = m_tEnterSize.Right - m_tEnterSize.Left
         lHeight = m_tEnterSize.Bottom - m_tEnterSize.Top
      End If
   End If
End Sub

Private Sub UserControl_Initialize()
   AllowTearOff = True
   MoveOnFormMouseDown = True
   ' Minimum number of pixels moved before
   ' drop down item can be dragged off:
   m_iHysteresis = 8
   ' use get sys metrics here
   Set m_cNCM = New cNCMetrics
   m_lTearOffHeight = 6
   m_bAllowResize = True
End Sub

Private Sub UserControl_InitProperties()
   
   ' Set ambient properties:
   UserControl.Extender.Align = 1
   UserControl.Extender.ToolTipText = "Drag to make this menu float"
   pInitialise
   
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tP As POINTAPI

   If (Button And vbLeftButton) = vbLeftButton Then
      
      m_bMouseDown = True
      GetCursorPos m_tPMouseDown
      
      ' Move the form if we are floating:
      If m_eShowState = ewssFloating Then
         LSet tP = m_tPMouseDown
         ScreenToClient m_hWnd, tP
         If PtInRect(m_tButtonR, tP.x, tP.y) <> 0 Then
            ' Over close button:
            m_bButtonDown = True
            m_bButtonOver = True
            pDrawCloseButton UserControl.hDC
         Else
            ' Over caption:
            ReleaseCapture
            SendMessageLong ParenthWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
         End If
      End If
      
   Else
      RaiseEvent RightClick(x, y)
   End If
   
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tR As RECT
Dim tP As POINTAPI
Dim bRelease As Boolean
   
   If m_bMouseDown Then
      GetCursorPos tP
      If m_eShowState = ewssDropped Then
         ' Inside the float box?
         GetWindowRect UserControl.hwnd, tR
         If PtInRect(tR, tP.x, tP.y) = 0 Then
            bRelease = (Abs(tP.y - m_tPMouseDown.y) > m_iHysteresis) Or _
               (tP.x <= tR.Left - m_iHysteresis Or tP.y >= tR.Right + m_iHysteresis)
            If bRelease Then
               ' release the window!
               ShowState = ewssFloating
               ReleaseCapture
               SendMessageLong ParenthWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
               If (tP.x <= tR.Left - m_iHysteresis Or tP.y >= tR.Right + m_iHysteresis) Then
                  ' Move the window so the cursor is centered:
                  MoveWindow ParenthWnd, tR.Left + (tR.Right - tR.Left) \ 2, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, 1
               End If
            End If
         End If
      ElseIf m_eShowState = ewssFloating Then
         If m_bButtonDown Then
            ScreenToClient m_hWnd, tP
            If PtInRect(m_tButtonR, tP.x, tP.y) <> 0 Then
               If Not m_bButtonOver Then
                  m_bButtonOver = True
                  pDrawCloseButton UserControl.hDC
               End If
            Else
               If m_bButtonOver Then
                  m_bButtonOver = False
                  pDrawCloseButton UserControl.hDC
               End If
            End If
         End If
      End If
   Else
      'Debug.Print "MouseOver:"; m_bMouseOver
      If Not m_bMouseOver Then
         m_cMT.StartMouseTracking
         m_bMouseOver = True
         DrawTitleBar
      End If
   End If
   
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tP As POINTAPI
   If m_bMouseDown And m_bButtonDown And m_eShowState = ewssFloating Then
      m_bButtonDown = False
      m_bButtonOver = False
      pDrawCloseButton UserControl.hDC
      GetCursorPos tP
      ScreenToClient m_hWnd, tP
      If PtInRect(m_tButtonR, tP.x, tP.y) <> 0 Then
         ' we clicked the close button
         PostMessage m_hWnd, WM_CLOSE, 0, 0
      End If
   End If
   m_bMouseDown = False
End Sub

Private Sub UserControl_Paint()
   DrawTitleBar
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   UserControl.Extender.ToolTipText = "Drag to make this menu float"
   MoveOnFormMouseDown = PropBag.ReadProperty("MoveOnFormMouseDown", True)
   pInitialise
   AllowTearOff = PropBag.ReadProperty("AllowTearOff", True)
   AllowResize = PropBag.ReadProperty("AllowResize", True)
   Caption = PropBag.ReadProperty("Caption", "")
End Sub

Private Sub UserControl_Resize()
   If AllowTearOff Then
      DrawTitleBar
   End If
   RaiseEvent CaptionResize
End Sub

Private Sub UserControl_Terminate()
   pTerminate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   PropBag.WriteProperty "AllowTearOff", AllowTearOff, True
   PropBag.WriteProperty "AllowResize", AllowResize, True
   PropBag.WriteProperty "MoveOnFormMouseDown", MoveOnFormMouseDown, True
   PropBag.WriteProperty "Caption", Caption, ""
End Sub
