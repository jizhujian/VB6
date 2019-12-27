VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cNCMetrics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SPI_GETICONMETRICS = 45
Private Const SPI_GETICONTITLELOGFONT = 31
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64

' Normal log font structure:
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
Private Enum CNCMetricsFontWeightConstants
   FW_DONTCARE = 0
   FW_THIN = 100
   FW_EXTRALIGHT = 200
   FW_ULTRALIGHT = 200
   FW_LIGHT = 300
   FW_NORMAL = 400
   FW_REGULAR = 400
   FW_MEDIUM = 500
   FW_SEMIBOLD = 600
   FW_DEMIBOLD = 600
   FW_BOLD = 700
   FW_EXTRABOLD = 800
   FW_ULTRABOLD = 800
   FW_HEAVY = 900
   FW_BLACK = 900
End Enum
' For some bizarre reason, maybe to do with byte
' alignment, the LOGFONT structure we must apply
' to NONCLIENTMETRICS seems to require an LF_FACESIZE
' 4 bytes smaller than normal:
Private Type NMLOGFONT
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
    lfFaceName(LF_FACESIZE - 4) As Byte
End Type
Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As NMLOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As NMLOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As NMLOGFONT
    lfStatusFont As NMLOGFONT
    lfMessageFont As NMLOGFONT
End Type
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
    ByVal uAction As Long, _
    ByVal uParam As Long, _
    lpvParam As Any, _
    ByVal fuWinIni As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private m_tNCM As NONCLIENTMETRICS
Private m_tLF As LOGFONT

Public Enum CNCMetricsFontTypes
    IconFont = 1
    CaptionFont = 2
    SMCaptionFont = 3
    MenuFOnt = 4
    StatusFont = 5
    MessageFont = 6
End Enum

Public Function GetMetrics() As Boolean
Dim lR As Long
    ' Get Non-client metrics:
    m_tNCM.cbSize = 340 'LenB(m_tNCM) - why doesn't this go?
    lR = SystemParametersInfo( _
            SPI_GETNONCLIENTMETRICS, _
            0, _
            m_tNCM, _
            0)
    If (lR <> 0) Then
        ' Get icon font:
        lR = SystemParametersInfo( _
            SPI_GETICONTITLELOGFONT, _
            0, _
            m_tLF, _
            0)
        GetMetrics = (lR <> 0)
    End If
End Function
Property Get Font( _
      ByVal hDC As Long, _
      ByVal eFontNum As CNCMetricsFontTypes _
   ) As Long
Dim lR As Long
Dim tLF As LOGFONT

   Select Case eFontNum
   Case StatusFont
      CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
   Case SMCaptionFont
      CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
   Case MessageFont
      CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
   Case MenuFOnt
      CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
   Case IconFont
      CopyMemory tLF, m_tLF, LenB(m_tLF)
   Case CaptionFont
      CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
   Case Else
      Exit Property
   End Select
    
   ' If you want an API hFont, just do this:
   Font = CreateFontIndirect(tLF)
   
   ' This demonstrates how to return a VB style font.
   ' Remember to use DeleteObject hFont when you've
   ' finished with it.
   'Dim sFnt As New StdFont
   'pLogFontToStdFont tLF, hDC, sFnt
   'Set Font = sFnt
   
End Property
Private Sub pLogFontToStdFont(ByRef tLF As LOGFONT, ByVal hDC As Long, ByRef sFnt As StdFont)
   With sFnt
      .Name = StrConv(tLF.lfFaceName, vbUnicode)
      If tLF.lfHeight < 1 Then
         .Size = Abs((72# / GetDeviceCaps(hDC, LOGPIXELSY)) * tLF.lfHeight)
      Else
         .Size = tLF.lfHeight
      End If
      .Charset = tLF.lfCharSet
      .Italic = Not (tLF.lfItalic = 0)
      .Underline = Not (tLF.lfUnderline = 0)
      .Strikethrough = Not (tLF.lfStrikeOut = 0)
      .Bold = (tLF.lfWeight > FW_REGULAR)
   End With
End Sub
Property Get CaptionHeight() As Long
    CaptionHeight = m_tNCM.iCaptionHeight
End Property
Property Get CaptionWIdth() As Long
    CaptionWIdth = m_tNCM.iCaptionWidth
End Property
Property Get MenuHeight() As Long
    MenuHeight = m_tNCM.iMenuHeight
End Property
Property Get MenuWidth() As Long
    MenuWidth = m_tNCM.iMenuWidth
End Property
Property Get ScrollHeight() As Long
    ScrollHeight = m_tNCM.iScrollHeight
End Property
Property Get ScrollWidth() As Long
    ScrollWidth = m_tNCM.iScrollWidth
End Property
Property Get SMCaptionHeight() As Long
    SMCaptionHeight = m_tNCM.iSMCaptionHeight
End Property
Property Get SMCaptionWIdth() As Long
    SMCaptionWIdth = m_tNCM.iSMCaptionWidth
End Property
Property Get BorderWidth() As Long
    BorderWidth = m_tNCM.iBorderWidth
End Property
