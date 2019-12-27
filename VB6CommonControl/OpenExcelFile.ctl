VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.UserControl OpenExcelFile 
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7950
   ScaleHeight     =   1815
   ScaleWidth      =   7950
   ToolboxBitmap   =   "OpenExcelFile.ctx":0000
   Begin VB.CommandButton cmOK 
      Caption         =   "打开EXCEL表文件"
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox cboFilterFieldValue 
      Height          =   300
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1080
      Width           =   4455
   End
   Begin VB.ComboBox cboFilterFieldName 
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cboSheet 
      Height          =   300
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   720
      Width           =   7215
   End
   Begin VB.ComboBox cboDriver 
      Height          =   300
      ItemData        =   "OpenExcelFile.ctx":0312
      Left            =   720
      List            =   "OpenExcelFile.ctx":031C
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1935
   End
   Begin VB6CommonControl.TextBoxEx txtFileName 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      ShowLookupButton=   -1  'True
   End
   Begin MSComDlg.CommonDialog cdOpenFile 
      Left            =   7080
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Excel 97-2003 工作薄 (*.xls)|*.xls|Excel 工作薄 (*.xlsx)|*.xlsx"
   End
   Begin VB.Label lblEqual 
      AutoSize        =   -1  'True
      Caption         =   "＝"
      Height          =   180
      Left            =   3240
      TabIndex        =   8
      Top             =   1140
      Width           =   180
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "过滤："
      Height          =   180
      Left            =   0
      TabIndex        =   6
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lblSheet 
      AutoSize        =   -1  'True
      Caption         =   "工作表："
      Height          =   180
      Left            =   0
      TabIndex        =   4
      Top             =   780
      Width           =   720
   End
   Begin VB.Label lblDriver 
      AutoSize        =   -1  'True
      Caption         =   "格式："
      Height          =   180
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   540
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "文件名："
      Height          =   180
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   720
   End
End
Attribute VB_Name = "OpenExcelFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mstrConnectionString As String
Private mstrWhereClause As String
Private mstrLastFileName As String
Private mblnRebuildConnectionString As Boolean
Private macFileSystem As AutoComplete
Private mblnLockEvent As Boolean
Private mblnDisableErrMsg As Boolean

Public Event Change()
Public Event Execute()

Public Property Get ConnectionString() As String
  ConnectionString = mstrConnectionString
End Property

Public Property Get FromClause() As String
  If (cboSheet.ListIndex >= 0) Then
    FromClause = cboSheet.Text
  End If
End Property

Public Property Get WhereClause() As String
  WhereClause = mstrWhereClause
End Property

Private Sub cmOK_Click()
  If (Trim(txtFileName.Text) = "") Then
    MsgBox "请选择EXCEL文件。", vbExclamation, "打开EXCEL表文件"
    txtFileName.SetFocus
    Exit Sub
  End If
  If (cboSheet.ListIndex = -1) Then
    MsgBox "请选择工作表。", vbExclamation, "打开EXCEL表文件"
    cboSheet.SetFocus
    Exit Sub
  End If
  RaiseEvent Execute
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
  macFileSystem.DroppedDown
End Sub

Private Sub UserControl_Initialize()
  Set macFileSystem = New AutoComplete
  macFileSystem.Options = AutoCompleteOptionSuggest + AutoCompleteOptionUpDownKeyDropsList + AutoCompleteOptionUseTab
  macFileSystem.Init txtFileName.SubHWnd, AutoCompleteSourceFileSystem
  mblnLockEvent = True
  cboDriver.ListIndex = 0
  mblnLockEvent = False
End Sub

Private Sub Reset()
  mblnRebuildConnectionString = True
  mstrConnectionString = ""
  cboSheet.Clear
  ResetWhereClause
End Sub

Private Sub ResetWhereClause()
  mstrWhereClause = ""
  cboFilterFieldName.Clear
  cboFilterFieldValue.Clear
End Sub

Private Sub txtFileName_Change()
  txtFileName.Tag = "Changed"
  RaiseEvent Change
End Sub

Private Sub txtFileName_LookupButtonClick()
  With cdOpenFile
    .FileName = Trim(txtFileName.Text)
    .ShowOpen
    If .FileName > "" Then
      txtFileName.Text = .FileName
    End If
  End With
End Sub

Private Sub txtFileName_Validate(Cancel As Boolean)
  Dim objFile As dotNET2COM.file
  Dim strFileName As String
  If (txtFileName.Tag = "Changed") Then
    strFileName = Trim(txtFileName.Text)
    If (mstrLastFileName <> strFileName) Then
      If (strFileName > "") Then
        Set objFile = New dotNET2COM.file
        If objFile.ExistsFile(strFileName) = False Then
          Cancel = True
          If (mblnDisableErrMsg = False) Then
            MsgBox "EXCEL文件不存在。", vbCritical, "打开EXCEL表文件"
          End If
          Exit Sub
        End If
        If objFile.HasExtension(strFileName) Then
          If (LCase(objFile.GetExtension(strFileName)) = ".xlsx") And (cboDriver.ListIndex = 0) Then
            cboDriver.ListIndex = 1
          End If
        End If
      End If
      mstrLastFileName = strFileName
      Reset
    End If
    txtFileName.Tag = ""
  End If
End Sub

Private Sub cboDriver_Click()
  Reset
  If Not mblnLockEvent Then
    RaiseEvent Change
  End If
End Sub

Private Sub cboDriver_Validate(Cancel As Boolean)
  Dim objFile As dotNET2COM.file
  Dim strFileName As String
  If cboDriver.ListIndex > 0 Then Exit Sub
  strFileName = Trim(txtFileName.Text)
  If (strFileName > "") Then
    Set objFile = New dotNET2COM.file
    If objFile.HasExtension(strFileName) Then
      If (LCase(objFile.GetExtension(strFileName)) = ".xlsx") Then
        MsgBox ".xlsx文件请选择Excel 2007以上格式。", vbCritical, "打开EXCEL表文件"
        Cancel = True
        Exit Sub
      End If
    End If
  End If
End Sub

Private Sub cboDriver_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub cboSheet_GotFocus()
  Dim strConnectionString As String
  Dim strFileName As String
  Dim cnn As ADODB.Connection
  Dim rs As ADODB.Recordset
  If mblnRebuildConnectionString = False Then
    Exit Sub
  End If
  strFileName = Trim(txtFileName.Text)
  If (strFileName = "") Then
    Exit Sub
  End If
  Select Case cboDriver.ListIndex
    Case 0
      strConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strFileName & ";Extended Properties=""Excel 8.0;IMEX=1"""
    Case 1
      strConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strFileName & ";Extended Properties=""Excel 12.0;IMEX=1"""
    Case 2
      strConnectionString = "Provider=Microsoft.ACE.OLEDB.15.0;Data Source=" & strFileName & ";Extended Properties=""Excel 12.0;IMEX=1"""
    Case 3
      strConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & strFileName & ";Extended Properties=""Excel 12.0;IMEX=1"""
  End Select
  Set cnn = New ADODB.Connection
  On Error Resume Next
  Err.Clear
  cnn.Open strConnectionString
  If (cnn.State = 1) Then
    mblnRebuildConnectionString = False
    mstrConnectionString = strConnectionString
    Set rs = cnn.OpenSchema(adSchemaTables)
    Do Until rs.EOF
      If Right(rs("TABLE_NAME").Value, 1) = "$" Then
        cboSheet.AddItem rs("TABLE_NAME").Value
      End If
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cnn.Close
  End If
  Set cnn = Nothing
  If Err.Number <> 0 And mblnDisableErrMsg = False Then
    MsgBox Err.Source & vbCrLf & Err.Description, vbCritical, "打开EXCEL表文件"
  End If
End Sub

Private Sub cboSheet_Click()
  Dim cnn As ADODB.Connection
  Dim rs As ADODB.Recordset
  On Error GoTo HERROR
  RaiseEvent Change
  ResetWhereClause
  Set cnn = New ADODB.Connection
  cnn.Open mstrConnectionString
  Set rs = cnn.OpenSchema(adSchemaColumns, Array(Empty, Empty, cboSheet.Text, Empty))
  Do Until rs.EOF
    cboFilterFieldName.AddItem rs("COLUMN_NAME").Value
    cboFilterFieldName.ItemData(cboFilterFieldName.ListCount - 1) = rs("DATA_TYPE").Value
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  cnn.Close
  Set cnn = Nothing
  If cboFilterFieldName.ListCount > 0 Then
    cboFilterFieldName.AddItem "", 0
  End If
  Exit Sub
HERROR:
  If Not cnn Is Nothing Then
    If cnn.State = 1 Then
      cnn.Close
    End If
  End If
  MsgBox Err.Source & vbCrLf & Err.Description, vbCritical, "打开EXCEL表文件"
End Sub

Private Sub cboSheet_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub cboFilterFieldName_Click()
  Dim cnn As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim sql As String
  mstrWhereClause = ""
  On Error GoTo HERROR
  RaiseEvent Change
  cboFilterFieldValue.Clear
  If cboFilterFieldName.ListIndex <= 0 Then
    Exit Sub
  End If
  Set cnn = New ADODB.Connection
  cnn.Open mstrConnectionString
  sql = "SELECT DISTINCT `" & cboFilterFieldName.Text & "` AS FFilterFieldValue" & vbCrLf & _
    "FROM `" & cboSheet.Text & "`" & vbCrLf & _
    "WHERE (`" & cboFilterFieldName.Text & "` IS NOT NULL)" & vbCrLf & _
    "ORDER BY `" & cboFilterFieldName.Text & "`"
  Set rs = cnn.Execute(sql, , adCmdText)
  Do Until rs.EOF
    cboFilterFieldValue.AddItem rs("FFilterFieldValue").Value
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  cnn.Close
  Set cnn = Nothing
  Exit Sub

HERROR:
  If Not cnn Is Nothing Then
    If cnn.State = 1 Then
      cnn.Close
    End If
  End If
  MsgBox Err.Source & vbCrLf & Err.Description, vbCritical, "打开EXCEL表文件"
End Sub

Private Sub cboFilterFieldName_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub cboFilterFieldValue_Click()
  Dim s As String
  Dim v As String
  RaiseEvent Change
  v = cboFilterFieldValue.Text
  Select Case cboFilterFieldName.ItemData(cboFilterFieldName.ListIndex)
    Case adWChar
      s = "'"
      v = Replace(v, "'", "''")
    Case adDate
      s = "'"
  End Select
  mstrWhereClause = "(`" & cboFilterFieldName.Text & "` = " & s & v & s & ")"
End Sub

Private Sub cboFilterFieldValue_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
    KeyCode = 0
    SendKeys "{TAB}"
  End If
End Sub

Private Sub UserControl_Resize()
  Dim dblWidth As Double
  On Error Resume Next
  dblWidth = UserControl.ScaleWidth - txtFileName.Left
  If dblWidth < 1000 Then Exit Sub
  txtFileName.Width = dblWidth
  cboSheet.Width = dblWidth
  cboFilterFieldName.Width = dblWidth / 2.5
  lblEqual.Left = cboFilterFieldName.Left + cboFilterFieldName.Width
  cboFilterFieldValue.Left = lblEqual.Left + lblEqual.Width
  cboFilterFieldValue.Width = UserControl.ScaleWidth - cboFilterFieldValue.Left
End Sub

Public Sub LoadDefaultValue(ByVal col As Collection)
  Dim blnCancel As Boolean
  On Error GoTo HERROR
  mblnDisableErrMsg = True
  txtFileName.Text = col("FileName")
  cboDriver.ListIndex = col("Driver")
  txtFileName_Validate blnCancel
  If txtFileName.Text = "" Or blnCancel Then
    mblnDisableErrMsg = False
    Exit Sub
  End If
  cboSheet_GotFocus
  SelectListItemByText cboSheet, col("Sheet")
  SelectListItemByText cboFilterFieldName, col("FilterFieldName")
  SelectListItemByText cboFilterFieldValue, col("FilterFieldValue")
HERROR:
  mblnDisableErrMsg = False
End Sub

Public Function SaveDefaultValue() As Collection
  Dim col As New Collection
  col.Add Trim(txtFileName.Text), "FileName"
  col.Add cboDriver.ListIndex, "Driver"
  col.Add cboSheet.Text, "Sheet"
  col.Add cboFilterFieldName.Text, "FilterFieldName"
  col.Add cboFilterFieldValue.Text, "FilterFieldValue"
  Set SaveDefaultValue = col
End Function
