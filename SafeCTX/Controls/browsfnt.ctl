VERSION 5.00
Begin VB.UserControl FontBrowser 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   KeyPreview      =   -1  'True
   ScaleHeight     =   465
   ScaleWidth      =   2535
   ToolboxBitmap   =   "browsfnt.ctx":0000
   Begin VB.PictureBox picBrowse 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1845
      ScaleHeight     =   300
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   135
      Width           =   330
   End
   Begin VB.Label lblFont 
      BackColor       =   &H80000005&
      Height          =   225
      Left            =   270
      TabIndex        =   0
      Top             =   135
      Width           =   1395
   End
End
Attribute VB_Name = "FontBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Browse for font control from Paul Duffield's Visual Basic Resources.
'You are free to use this source code as is, or modified in your own
'software provided that you agree that Paul Duffield has no
'responsibility or liability whatsoever for any loss or damage
'occasioned by its use.
'
'Some portions of this code have been modified from an example found
'on the internet.  Original author unknown.
'
'To modify the font selection behaviour of this control, see the code
'in the 'picBrowse_Click' and 'SelectFont' procedures.

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const CF_Apply = &H200
Private Const CF_EnableHook = &H8
Private Const CF_EnableTemplate = &H10
Private Const CF_EnableTemplateHandle = &H20
Private Const CF_FontNotSupported = &H238
Private Const CF_ScreenFonts = &H1
Private Const CF_PrinterFonts = &H2
Private Const CF_BOTH = &H3
Private Const CF_EFFECTS = &H100
Private Const LF_FACESIZE = 32
Private Const Bold_FontType = &H100
Private Const Italic_FontType = &H200
Private Const Regular_FontType = &H400
Private Const CF_InitToLogFontStruct = &H40

Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type TChooseFont
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type

Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function ChooseFont Lib "COMDLG32" Alias "ChooseFontA" (chfont As TChooseFont) As Long

'DrawEdge constants
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

' Border flags
Private Const BF_LEFT As Long = &H1
Private Const BF_TOP As Long = &H2
Private Const BF_RIGHT As Long = &H4
Private Const BF_BOTTOM As Long = &H8
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)

'System Metrics
Private Const SM_CXVSCROLL = 2 'width of scroll bar and combo buttons

'Event Declarations:
Event Change(Font As StdFont) 'MappingInfo=lblFont,lblFont,-1,Change

Public Sub AboutBox()
    About
End Sub

Private Sub picBrowse_Click()
    
    Dim curFont As Font
    Dim lColor As Long
    
    Set curFont = lblFont.Font
    
    'NOTE: If you are going to use system colors you will need to disable color
    'selection by removing 'lColor' from the call to SelectFont, and commenting
    'the line 'UserControl.ForeColor = lColor'
    lColor = UserControl.ForeColor
    
    'Show the browse for file dialog
    If SelectFont(curFont, , UserControl.Extender.Parent.hWnd, lColor) Then
        Set lblFont.Font = curFont
        UserControl.ForeColor = lColor
        lblFont.Caption = lblFont.Font.Name
        
        RaiseEvent Change(curFont)
    End If
    
End Sub

Private Sub picBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PaintAsButton picBrowse, True
End Sub

Private Sub picBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PaintAsButton picBrowse
End Sub

Private Sub picBrowse_Paint()
    
    PaintAsButton picBrowse
    
End Sub

Private Sub UserControl_EnterFocus()
    UserControl_Paint
End Sub

Private Sub UserControl_ExitFocus()
    UserControl_Paint
End Sub

Private Sub UserControl_Initialize()
    'Set the width of the 'button'
    picBrowse.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    'Keypreview is set, so we get all of the keypresses here first.
    'Check for keypresses which should cause the Browse dialog to show
    'Alt and down arrow.
    If KeyCode = vbKeyDown And Shift = 4 Then
        picBrowse_Click
    End If
End Sub

Private Sub UserControl_Paint()
    
    Dim rct As RECT
   
   'Draw a focus rectangle and set label color as necessary
    If GetFocus = picBrowse.hWnd Then 'the only control capable of receiving focus on the UserControl
        GetClientRect UserControl.hWnd, rct
        With rct
            .Left = .Left + 1
            .Right = .Right - ((picBrowse.Width / Screen.TwipsPerPixelX) + 1)
            .Top = .Top + 1
            .Bottom = .Bottom - 1
        End With
        DrawFocusRect UserControl.hdc, rct
        With lblFont
            .ForeColor = vbHighlightText
            .BackColor = vbHighlight
        End With
    Else
        UserControl.Cls
        With lblFont
            .ForeColor = UserControl.ForeColor
            .BackColor = vbWindowBackground
        End With
    
    End If
    
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    'Position the constituent controls
    picBrowse.Move UserControl.ScaleWidth - picBrowse.Width, 0, picBrowse.Width, UserControl.ScaleHeight
    lblFont.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, UserControl.ScaleWidth - (picBrowse.Width + (4 * Screen.TwipsPerPixelX)), UserControl.ScaleHeight - (4 * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Show()
    'Get the tooltip
    lblFont.ToolTipText = UserControl.Extender.ToolTipText
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFont,lblFont,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFont,lblFont,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFont,lblFont,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFont,lblFont,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblFont.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblFont.Font = New_Font
    PropertyChanged "Font"
    lblFont.Caption = lblFont.Font.Name
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblFont,lblFont,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    lblFont.Refresh
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblFont.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblFont.Caption = lblFont.Font.Name
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblFont.Font, Ambient.Font)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set lblFont.Font = Ambient.Font
    lblFont.Caption = lblFont.Font.Name
End Sub

Private Sub PaintAsButton(objControl As Object, Optional bPressed As Boolean = False)
    
    Dim rct As RECT
    
    'Draw a 3D sunken border
    GetClientRect picBrowse.hWnd, rct
    With objControl
        .Cls
        If bPressed Then
            DrawEdge .hdc, rct, BDR_SUNKENOUTER, BF_RECT
        Else
            DrawEdge .hdc, rct, BDR_RAISED, BF_RECT
        End If
        .Font.Name = "System"
        .Font.Size = "10"
        .Font.Bold = True
        .CurrentX = ((.ScaleWidth / 2) - (.TextWidth("...") / 2)) - (bPressed * Screen.TwipsPerPixelX)
        .CurrentY = .ScaleHeight - .TextHeight("...") - (2 * Screen.TwipsPerPixelY) - (bPressed * Screen.TwipsPerPixelY)
        picBrowse.Print "..."
    End With

End Sub

'Credits:
'The code in the following procedures was modified from an example downloaded
'from the internet.  Original Author unknown.

Private Function SelectFont(curFont As Font, Optional PrinterDC As Long = -1, Optional Owner As Long = -1, Optional Color As Long = vbBlack) As Boolean

    Dim m_lApiReturn As Long
    Dim m_lExtendedError As Long
    Dim flags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PrinterFonts Then PrinterDC = Printer.hdc
    Else
        flags = flags Or CF_PrinterFonts
    End If
    ' Must have some fonts
    If (flags And CF_PrinterFonts) = 0 Then flags = flags Or CF_ScreenFonts
    ' Color can take initial color, receive chosen color
    'If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    flags = flags Or CF_EFFECTS
    
    ' Put in required internal flags and remove unsupported
    flags = (flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(curFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = curFont.Weight
    fnt.lfItalic = curFont.Italic
    fnt.lfUnderline = curFont.Underline
    fnt.lfStrikeOut = curFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, curFont.Name

    ' Initialize TSelectFont variable
    Dim cf As TChooseFont
    With cf
        .lStructSize = Len(cf)
        If Owner <> -1 Then .hWndOwner = Owner
        .hdc = PrinterDC
        .lpLogFont = VarPtr(fnt)
        .iPointSize = curFont.Size * 10
        .flags = flags
        .rgbColors = Color
        ' All other fields zero
    End With
    m_lApiReturn = ChooseFont(cf)
    Select Case m_lApiReturn
    Case 1
        ' Success
        SelectFont = True
        flags = cf.flags
        Color = cf.rgbColors
        With curFont
            .Bold = cf.nFontType And Bold_FontType
            .Italic = fnt.lfItalic
            .Strikethrough = fnt.lfStrikeOut
            .Underline = fnt.lfUnderline
            .Weight = fnt.lfWeight
            .Size = cf.iPointSize / 10
            .Name = BytesToStr(fnt.lfFaceName)
        End With
    Case 0
        ' Cancelled
        SelectFont = False
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        SelectFont = False
    End Select
        
End Function

Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        CopyMemoryStr ab(LBound(ab)), s, cab
    End If
End Sub

Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function

