VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TrueDBGridEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngTrueDBGridObjPtr As Long
Private mobjRedStyle As Object 'Style
Private mblnExistEmplyColumn As Boolean

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub Init(ByVal tdbgrd As Object)
  If TypeName(tdbgrd) = "TDBGrid" Then
    mlngTrueDBGridObjPtr = ObjPtr(tdbgrd)
    SetGeneral
    SetStyles
'    SetPrintInfo
    mblnExistEmplyColumn = True
  End If
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get TrueDBGridControl() As Object
  Dim objPtr2Object As New Ptr2Object
  Set TrueDBGridControl = objPtr2Object.ObjectFromPtr(mlngTrueDBGridObjPtr)
  Set objPtr2Object = Nothing
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Sub SetGeneral()
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  With tdbgrd
    .AllowAddNew = False
    .AllowUpdate = False
    .AllowDelete = False
    .AllowColSelect = False
    .MarqueeStyle = 0 'dbgDottedCellBorder
    .HeadLines = 2
    .FootLines = 1.2
    .RowHeight = 315
    .AnchorRightColumn = True
    .TabAcrossSplits = True
    .FetchRowStyle = True
    .WrapCellPointer = True
    .Appearance = 3 'dbgXPTheme
  End With
End Sub

' ***********************************************************************************************
'设置数据网格的列数
' ***********************************************************************************************
Public Property Let Cols(ByVal intValue As Integer)
  Dim tdbgrd As Object
  Dim intCounter As Integer
  Dim colTemp As Object
  Set tdbgrd = TrueDBGridControl
  With tdbgrd.Columns
    '删除多余的列
    While .count > intValue
      .Remove .count - 1
    Wend
    '增加指定的列数
    For intCounter = .count To intValue - 1
      Set colTemp = .Add(intCounter)
      colTemp.Visible = True
    Next
  End With
End Property

' ***********************************************************************************************
'预置数据网格样式
' ***********************************************************************************************
Private Sub SetStyles()
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  On Error Resume Next
  With tdbgrd.Styles
    '通用表体样式：白/黑
    With .Item("Normal")
      .VerticalAlignment = 2 'dbgVertCenter
      .WrapText = True
    End With
    '列标题样式：浅灰/黑
    With .Item("Heading")
      .Alignment = 2 'dbgCenter
      .WrapText = True
    End With
    '列脚注样式：浅灰/黑
    With .Item("Footing")
      .WrapText = True
    End With
    '网格标题样式：浅灰/黑，粗体，大字
    With .Item("Caption")
      .Font.Name = "楷体_GB2312"
      .Font.Bold = True
      .Font.Size = 18
      .Alignment = 2 'dbgCenter
      .WrapText = True
    End With
    '选定行或列样式：蓝/白
    With .Item("Selected")
      .BackColor = vbBlue
      .ForeColor = vbWhite
      .WrapText = True
    End With
    '高亮度行样式：黑/白
    With .Item("HighlightRow")
      .BackColor = vbBlack
      .ForeColor = vbWhite
      .WrapText = True
    End With
    '奇数行样式：白/黑
    With .Item("OddRow")
      .WrapText = True
    End With
    '偶数行样式：浅灰/黑
    With .Item("EvenRow")
      .BackColor = &H80000016
      .WrapText = True
    End With
    '红字行样式：白/红
    With .Add("RedOddRow")
      .ForeColor = vbRed
      .WrapText = True
    End With
    '红字行样式：浅灰/红
    With .Add("RedEvenRow")
      .BackColor = &H80000016
      .ForeColor = vbRed
      .WrapText = True
    End With
    '奇数行样式：青/黑
    With .Add("CyanRow")
      .BackColor = vbCyan
      .WrapText = True
    End With
    '红字行样式：青/红
    With .Add("RedCyanRow")
      .BackColor = vbCyan
      .ForeColor = vbRed
      .WrapText = True
    End With
    '工具提示样式：浅黄/黑
    With .Add("Tip")
      .BackColor = vbInfoBackground
      .ForeColor = vbInfoText
      .WrapText = True
    End With
    '被删除的记录样式：白/浅灰
    With .Add("Deleted")
      .ForeColor = vbButtonFace
      .WrapText = True
    End With
  End With
End Sub

' ***********************************************************************************************
'锁定网格左边不滚动列
' ***********************************************************************************************
Public Sub LockColumns(ByVal intLeftCols As Integer, Optional ByVal intRightCols As Integer)
  Dim tdbgrd As Object
  Dim intSplit As Integer
  Dim intSplits As Integer
  Dim intCol As Integer
  Dim splTemp As Object 'Split
  Dim intLeftInvisibleCols As Integer
  Dim intRightInvisibleCols As Integer
  Set tdbgrd = TrueDBGridControl
  If intLeftCols > 0 And intRightCols > 0 Then intSplits = 3 Else intSplits = 2
  '删除其余拆分
  While tdbgrd.Splits.count > 1
    tdbgrd.Splits.Remove tdbgrd.Splits.count - 1
  Wend
  With tdbgrd.Splits(0)
    .SizeMode = 0 'dbgScalable
    .Size = 1
  End With
  '增加拆分
  For intSplit = 1 To intSplits - 1
    Set splTemp = tdbgrd.Splits.Add(intSplit)
    splTemp.RecordSelectors = False
  Next
  '设置当前拆分
  tdbgrd.Split = 0
  '左边锁定列中如果包含隐藏列，现在将隐藏列加进锁定列中
  If intLeftCols > 0 Then
    For intCol = 0 To tdbgrd.Columns.count - 1
      If intCol < intLeftCols + intLeftInvisibleCols Then
        If tdbgrd.Columns(intCol).Visible = False Then
          intLeftInvisibleCols = intLeftInvisibleCols + 1
        End If
      Else
        Exit For
      End If
    Next
  End If
'右边锁定列中如果包含隐藏列，现在将隐藏列加进锁定列中
  If intRightCols > 0 Then
    For intCol = tdbgrd.Columns.count - 1 To 0 Step -1
      If intCol >= tdbgrd.Columns.count - (intRightCols + intRightInvisibleCols) Then
        If tdbgrd.Columns(intCol).Visible = False Then
          intRightInvisibleCols = intRightInvisibleCols + 1
        End If
      Else
        Exit For
      End If
    Next
  End If
'锁定左边列
  If intLeftCols > 0 Then
    For intSplit = 0 To tdbgrd.Splits.count - 1 - IIf(intRightCols > 0, 1, 0)
      For intCol = 0 To tdbgrd.Columns.count - 1
        With tdbgrd.Splits(intSplit).Columns(intCol)
          If intSplit = 0 Then
            If intCol >= intLeftCols + intLeftInvisibleCols Then
              .Visible = False
              .AllowSizing = False
            End If
          Else
            If intCol < intLeftCols + intLeftInvisibleCols Then
              .Visible = False
              .AllowSizing = False
            End If
          End If
        End With
      Next intCol
    Next intSplit
    tdbgrd.Splits(0).SizeMode = 2 'dbgNumberOfColumns
    tdbgrd.Splits(0).Size = intLeftCols
  End If
'锁定右边列
  If intRightCols > 0 Then
    For intSplit = tdbgrd.Splits.count - 1 To IIf(intLeftCols > 0, 1, 0) Step -1
      For intCol = tdbgrd.Columns.count - 1 To 0 Step -1
        With tdbgrd.Splits(intSplit).Columns(intCol)
          If intSplit = tdbgrd.Splits.count - 1 Then
            If intCol < tdbgrd.Columns.count - (intRightCols + intRightInvisibleCols) Then
              .Visible = False
              .AllowSizing = False
            End If
          Else
            If intCol >= tdbgrd.Columns.count - (intRightCols + intRightInvisibleCols) Then
              .Visible = False
              .AllowSizing = False
            End If
          End If
        End With
      Next intCol
    Next intSplit
    tdbgrd.Splits(tdbgrd.Splits.count - 1).SizeMode = 2 'dbgNumberOfColumns
    tdbgrd.Splits(tdbgrd.Splits.count - 1).Size = intRightCols
  End If
End Sub

' ***********************************************************************************************
'设置网格打印预览菜单
' ***********************************************************************************************
'Private Sub SetPrintInfo()
'  Dim tdbgrd As Object
'  Set tdbgrd = TrueDBGridControl
'  With tdbgrd.PrintInfo
'    .PreviewCaption = "打印预览"
'    .SetMenuText dbgpApsCaption, "调整页面设置"
'    .SetMenuText dbgpApsHorzCenter, "居中(&C)"
'    .SetMenuText dbgpApsHorzLeft, "左对齐(&L)"
'    .SetMenuText dbgpApsHorzRight, "右对齐(&R)"
'    .SetMenuText dbgpApsHorzText, "可打印区域的水平位置"
'    .SetMenuText dbgpApsMargins, "调整页边距，使可打印区域最大(&M)"
'    .SetMenuText dbgpApsOrientation, "选择最佳纸张方向(&O)"
'    .SetMenuText dbgpApsPageSetup, "页面设置(&P)"
'    .SetMenuText dbgpApsText, "纸张可打印区域太小，容纳不下打印内容。请检查/更改下面的设置，单击“确认”按钮调整页面设置并缩小打印内容。"
'    .SetMenuText dbgpApsVertBottom, "下对齐(&B)"
'    .SetMenuText dbgpApsVertCenter, "居中(&E)"
'    .SetMenuText dbgpApsVertText, "可打印区域的垂直位置"
'    .SetMenuText dbgpApsVertTop, "上对齐(&T)"
'    .SetMenuText dbgpDlgPagesCancel, "取消(&C)"
'    .SetMenuText dbgpDlgPagesCaption, "打印指定页"
'    .SetMenuText dbgpDlgPagesOk, "确认(&O)"
'    .SetMenuText dbgpDlgPagesPrompt, "请输入打印页号范围(-3,6,8-11,20-):"
'    .SetMenuText dbgpErrFileFormat, "文件格式错误。"
'    .SetMenuText dbgpErrFileOpen, "打开文件失败。"
'    .SetMenuText dbgpErrFileSave, "保存文件失败。"
'    .SetMenuText dbgpFilemaskAll, "所有文件"
'    .SetMenuText dbgpFilemaskArx, "Arx文件"
'    .SetMenuText dbgpMenuActualSize, "正常(&S)    Shift+Enter"
'    .SetMenuText dbgpMenuExit, "退出(&X)          Alt+F4"
'    .SetMenuText dbgpMenuFile, "文件(&F)"
'    .SetMenuText dbgpMenuFileClose, "关闭(&C)          Ctrl+F4"
'    .SetMenuText dbgpMenuFit, "整页(&W)    Enter"
'    .SetMenuText dbgpMenuFitWidth, "页宽(&D)    Ctrl+Enter"
'    .SetMenuText dbgpMenuOpen, "打开(&O)"
'    .SetMenuText dbgpMenuPageSetup, "页面设置..."
'    .SetMenuText dbgpMenuPgFirst, "首页(&F)    Home"
'    .SetMenuText dbgpMenuPgLast, "末页(&L)    End"
'    .SetMenuText dbgpMenuPgNext, "下页(&N)    PageDown"
'    .SetMenuText dbgpMenuPgPrev, "上页(&P)    PageUp"
'    .SetMenuText dbgpMenuPrint, "打印(&P)          Ctrl+P"
'    .SetMenuText dbgpMenuPrintCurrPage, "打印当前页(&C)    Ctrl+C"
'    .SetMenuText dbgpMenuPrintSomePages, "打印指定页(&S)    Ctrl+S"
'    .SetMenuText dbgpMenuSaveAs, "另存为(&A)"
'    .SetMenuText dbgpMenuView, "视图(&V)"
'    .SetMenuText dbgpMenuZoomIn, "放大(&I)    +"
'    .SetMenuText dbgpMenuZoomOut, "缩小(&O)    -"
'    .SetMenuText dbgpTipActualSize, "正常"
'    .SetMenuText dbgpTipFit, "整页"
'    .SetMenuText dbgpTipFitWidth, "页宽"
'    .SetMenuText dbgpTipOpen, "打开"
'    .SetMenuText dbgpTipPageOf, "页号"
'    .SetMenuText dbgpTipPageSetup, "页面设置"
'    .SetMenuText dbgpTipPgFirst, "首页"
'    .SetMenuText dbgpTipPgLast, "末页"
'    .SetMenuText dbgpTipPgNext, "下页"
'    .SetMenuText dbgpTipPgPrev, "上页"
'    .SetMenuText dbgpTipPrint, "打印"
'    .SetMenuText dbgpTipSaveAs, "另存为"
'    .SetMenuText dbgpTipStop, "停止处理"
'    .SetMenuText dbgpTipZoom, "缩放比例"
'    .SetMenuText dbgpTipZoomIn, "放大"
'    .SetMenuText dbgpTipZoomOut, "缩小"
'    .PageHeaderFont.Size = 20
'    .PageHeaderFont.Bold = True
'    .RepeatColumnHeaders = True
'    .PageFooter = "\t第\p页 共\P页"
'    .PreviewPageOf = " 第\p页 共\P页"
'  End With
'End Sub

Public Function AddColumn(ByVal strColumnTitle As String, ByVal strFieldName As String, ByVal dblNumChar As Double, _
  Optional ByVal intAlignment As AlignmentConstants, Optional ByVal blnVisible As Boolean = True, _
  Optional ByVal blnLocked As Boolean, Optional ByVal dblWrapText As Boolean) As Object
  Dim tdbgrd As Object
  Dim objColumn As Object
  Set tdbgrd = TrueDBGridControl
  If mblnExistEmplyColumn Then
    mblnExistEmplyColumn = False
    Set objColumn = tdbgrd.Columns(0)
  Else
    Set objColumn = tdbgrd.Columns.Add(tdbgrd.Columns.count)
  End If
  With objColumn
    .Caption = strColumnTitle
    .DataField = strFieldName
    .Width = dblNumChar * 200 + 50
    .Alignment = intAlignment
    .Visible = blnVisible
    .AllowSizing = blnVisible
    .Locked = blnLocked
    .WrapText = dblWrapText
  End With
  Set AddColumn = objColumn
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub AddNewRow()
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  With tdbgrd
    If .AllowAddNew = False Then Exit Sub
    If .DataChanged Then
      On Error Resume Next
      Err.Clear
      .Update
      If Err <> 0 Then Exit Sub
      On Error GoTo 0
    End If
    If .AddNewMode <> 1 Then 'dbgAddNewCurrent
      .MoveLast
      On Error Resume Next
      Err.Clear
      .Row = .Row + 1
      If Err <> 0 Then
        .FirstRow = .RowBookmark(1)
        .Row = .Row + 1
      End If
      On Error GoTo 0
      .col = 1
    End If
  End With
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub DeleteRow()
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  With tdbgrd
    If .AllowDelete = False Or .BOF Or .EOF Then Exit Sub
    .CurrentCellModified = False
    .EditActive = False
    .DataChanged = False
    If .AddNewMode = 1 Then Exit Sub 'dbgAddNewCurrent
    If IsNull(.Bookmark) = False Then .Delete
    If .AllowAddNew And .Row < 0 Then
      .Row = 0
      .col = 1
    End If
  End With
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Function RequiredTab() As Boolean
  Dim tdbgrd As Object
  Dim intIndex As Integer
  Set tdbgrd = TrueDBGridControl
  With tdbgrd
    If .FilterActive Then
      RequiredTab = False
    ElseIf .AllowAddNew Then
      If .AddNewMode = 1 And .Columns(.col).Locked = False Then RequiredTab = True 'dbgAddNewCurrent
    ElseIf IsNull(.Bookmark) Then
      RequiredTab = True
    ElseIf IsNull(.GetBookmark(1)) Then
      For intIndex = 0 To .Columns.count - 1
        If .Columns(intIndex).Visible And .Columns(intIndex).Order > .Columns(.col).Order Then Exit For
      Next
      If intIndex = .Columns.count Then RequiredTab = True
    End If
  End With
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
'Public Sub ExportToDelimitedFile()
'  Dim tdbgrd As TDBGrid
'  Set tdbgrd = TrueDBGridControl
'  tdbgrd.ExportToDelimitedFile Environ("Temp") & "\tdbgrd.txt", , vbTab, , , True
'  Shell "notepad.exe " & Environ("Temp") & "\tdbgrd.txt", vbMaximizedFocus
'End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Property Get RedStyle() As Object 'Style
  If (mobjRedStyle Is Nothing) Then
    Set mobjRedStyle = CreateObject("TrueOleDBGrid80.Style") ' New Style
    mobjRedStyle.ForeColor = vbRed
  End If
  Set RedStyle = mobjRedStyle
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub DisplayPositiveRed(ByVal col As Variant, Optional ByVal intSplit As Integer)
  Dim tdbgrd As Object
  Dim objColumn As Object
  Set tdbgrd = TrueDBGridControl
  Select Case TypeName(col)
    Case "String", "Byte", "Integer", "Long"
      Set objColumn = tdbgrd.Splits(intSplit).Columns(col)
      objColumn.AddRegexCellStyle -1, RedStyle, "^[0-9]" 'dbgAllCells
  End Select
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub DisplayNegativeRed(ByVal col As Variant, Optional ByVal intSplit As Integer)
  Dim tdbgrd As Object
  Dim objColumn As Object
  Set tdbgrd = TrueDBGridControl
  Select Case TypeName(col)
    Case "String", "Byte", "Integer", "Long"
      Set objColumn = tdbgrd.Splits(intSplit).Columns(col)
      objColumn.AddRegexCellStyle -1, RedStyle, "^-" 'dbgAllCells
  End Select
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function Update() As Boolean
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  If tdbgrd.DataChanged Then
    Err.Clear
    On Error Resume Next
    tdbgrd.Update
    On Error GoTo 0
    If Err <> 0 Then Exit Function
  End If
  Update = True
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub KeyDown(KeyCode As Integer, Shift As Integer)
  If RequiredTab Then
    If KeyCode = vbKeyReturn And Shift = 0 Then
      KeyCode = 0
      SendKeys "{TAB}"
    End If
  End If
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function FillArray(ByVal rs As ADODB.Recordset) As XArrayDB
  Dim arr As New XArrayDB
  Dim intCols As Integer
  Dim lngRows As Long
  Dim i As Long
  intCols = rs.Fields.count - 1
  arr.ReDim 1, 0, 0, intCols
  Do
    If rs.EOF Then Exit Do
    lngRows = lngRows + 1
    arr.AppendRows
    For i = 0 To intCols
      arr(lngRows, i) = rs.Fields(i).Value
    Next
    rs.MoveNext
  Loop
  rs.Close
  Set rs = Nothing
  Set FillArray = arr
End Function
