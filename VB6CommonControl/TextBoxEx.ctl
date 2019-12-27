VERSION 5.00
Begin VB.UserControl TextBoxEx 
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   ScaleHeight     =   6000
   ScaleWidth      =   8010
   ToolboxBitmap   =   "TextBoxEx.ctx":0000
   Begin VB.TextBox txtTextBox 
      Height          =   264
      Left            =   2160
      TabIndex        =   0
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Image imgLookup 
      Height          =   192
      Left            =   4560
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   192
   End
End
Attribute VB_Name = "TextBoxEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mblnShowLookupButton As Boolean
Private mblnKeyReturn2Tab As Boolean
Private mblnChanged As Boolean
Private mblnNumericValue As Boolean
Private mdblMinValue As Double
Private mdblMaxValue As Double
Private mintDecimalDigit As Integer

Public Event Change()
Public Event LookupButtonClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)

Public Property Get Text() As String
  Text = txtTextBox.Text
End Property

Public Property Let Text(ByVal Value As String)
  txtTextBox.Text = Value
  PropertyChanged "Text"
End Property

Public Property Get Enabled() As Boolean
  Enabled = txtTextBox.Enabled
End Property

Public Property Let Enabled(ByVal Value As Boolean)
  txtTextBox.Enabled = Value
  PropertyChanged "Enabled"
  UserControl_Resize
End Property

Public Property Get Locked() As Boolean
  Locked = txtTextBox.Locked
End Property

Public Property Let Locked(ByVal Value As Boolean)
  txtTextBox.Locked = Value
  PropertyChanged "Locked"
  UserControl_Resize
End Property

Public Property Get MaxLength() As Long
  MaxLength = txtTextBox.MaxLength
End Property

Public Property Let MaxLength(ByVal Value As Long)
  txtTextBox.MaxLength = Value
  PropertyChanged "MaxLength"
End Property

Public Property Get ShowLookupButton() As Boolean
  ShowLookupButton = mblnShowLookupButton
End Property

Public Property Let ShowLookupButton(ByVal Value As Boolean)
  mblnShowLookupButton = Value
  PropertyChanged "ShowLookupButton"
  UserControl_Resize
End Property

Public Property Get KeyReturn2Tab() As Boolean
  KeyReturn2Tab = mblnKeyReturn2Tab
End Property

Public Property Let KeyReturn2Tab(ByVal Value As Boolean)
  mblnKeyReturn2Tab = Value
  PropertyChanged "KeyReturn2Tab"
End Property

Public Property Get NumericValue() As Boolean
  NumericValue = mblnNumericValue
End Property

Public Property Let NumericValue(ByVal Value As Boolean)
  mblnNumericValue = Value
  PropertyChanged "NumericValue"
End Property

Public Property Get DecimalDigit() As Integer
  DecimalDigit = mintDecimalDigit
End Property

Public Property Let DecimalDigit(ByVal Value As Integer)
  mintDecimalDigit = Value
  PropertyChanged "DecimalDigit"
End Property

Public Property Get DoubleValue() As Double
  Dim strValue As String
  Dim dblValue As Double
  If NumericValue Then
    strValue = Trim(txtTextBox.Text)
    If strValue > "" Then
      If IsNumeric(strValue) Then
        If DecimalDigit = 0 Then
          dblValue = CLng(strValue)
        Else
          dblValue = CDbl(strValue)
          dblValue = Round(dblValue, DecimalDigit)
        End If
        DoubleValue = dblValue
      End If
    End If
  End If
End Property

Public Property Get MinValue() As Double
  MinValue = mdblMinValue
End Property

Public Property Let MinValue(ByVal Value As Double)
  mdblMinValue = Value
  PropertyChanged "MinValue"
End Property

Public Property Get MaxValue() As Double
  MaxValue = mdblMaxValue
End Property

Public Property Let MaxValue(ByVal Value As Double)
  mdblMaxValue = Value
  PropertyChanged "MaxValue"
End Property

Private Sub imgLookup_Click()
  RaiseEvent LookupButtonClick
End Sub

Private Sub txtTextBox_Change()
  mblnChanged = True
  RaiseEvent Change
End Sub

Private Sub txtTextBox_GotFocus()
  txtTextBox.SelStart = 0
  txtTextBox.SelLength = Len(txtTextBox.Text)
End Sub

Private Sub txtTextBox_KeyPress(KeyAscii As Integer)
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyDown(KeyCode, Shift)
  If KeyReturn2Tab And KeyCode = vbKeyReturn And Shift = 0 Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub txtTextBox_KeyUp(KeyCode As Integer, Shift As Integer)
  RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub txtTextBox_Validate(Cancel As Boolean)
  Dim strValue As String
  Dim dblValue As Double
  If mblnChanged Then
    If NumericValue Then
      strValue = Trim(txtTextBox.Text)
      If strValue > "" Then
        If Not IsNumeric(strValue) Then
          Cancel = True
          MsgBox "请输入数值。", vbCritical, "错误"
          Exit Sub
        End If
      End If
      dblValue = DoubleValue
      If (dblValue < MinValue) Then
        Cancel = True
        MsgBox "请重新输入数值。" & vbCrLf & "数值最小值：" & MinValue & "。", vbCritical, "错误"
        Exit Sub
      End If
      If (dblValue > MaxValue) Then
        Cancel = True
        MsgBox "请重新输入数值。" & vbCrLf & "数值最大值：" & MaxValue & "。", vbCritical, "错误"
        Exit Sub
      End If
    End If
    mblnChanged = False
  End If
End Sub

Private Sub UserControl_Initialize()
  mdblMinValue = -2147483648#
  mdblMaxValue = 2147483647
  Set imgLookup.Picture = LoadResIcon("search")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Text = .ReadProperty("Text", "")
    mblnChanged = False
    Enabled = .ReadProperty("Enabled", True)
    Locked = .ReadProperty("Locked", False)
    MaxLength = .ReadProperty("MaxLength", 0)
    ShowLookupButton = .ReadProperty("ShowLookupButton", False)
    KeyReturn2Tab = .ReadProperty("KeyReturn2Tab", True)
    NumericValue = .ReadProperty("NumericValue", False)
    DecimalDigit = .ReadProperty("DecimalDigit", 0)
    MinValue = .ReadProperty("MinValue", -2147483648#)
    MaxValue = .ReadProperty("MaxValue", 2147483647)
  End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    .WriteProperty "Text", Text, ""
    .WriteProperty "Enabled", Enabled, True
    .WriteProperty "Locked", Locked, False
    .WriteProperty "MaxLength", MaxLength, 0
    .WriteProperty "ShowLookupButton", ShowLookupButton, False
    .WriteProperty "KeyReturn2Tab", KeyReturn2Tab, True
    .WriteProperty "NumericValue", NumericValue, False
    .WriteProperty "DecimalDigit", DecimalDigit, 0
    .WriteProperty "MinValue", MinValue, -2147483648#
    .WriteProperty "MaxValue", MaxValue, 2147483647
  End With
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  If Not ShowLookupButton Or Not Enabled Or Locked Then
    txtTextBox.Move 0, 0, ScaleWidth, ScaleHeight
    imgLookup.Visible = False
  Else
    txtTextBox.Move 0, 0, ScaleWidth - imgLookup.Width, ScaleHeight
    imgLookup.Visible = True
    imgLookup.Move ScaleWidth - imgLookup.Width, (ScaleHeight - imgLookup.Height) / 2
  End If
End Sub

Public Property Get SubHWnd() As Long
  SubHWnd = txtTextBox.hWnd
End Property

Public Property Get Font() As StdFont
  Set Font = txtTextBox.Font
End Property

