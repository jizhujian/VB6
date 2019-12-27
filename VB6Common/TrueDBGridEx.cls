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
'�����������������
' ***********************************************************************************************
Public Property Let Cols(ByVal intValue As Integer)
  Dim tdbgrd As Object
  Dim intCounter As Integer
  Dim colTemp As Object
  Set tdbgrd = TrueDBGridControl
  With tdbgrd.Columns
    'ɾ���������
    While .count > intValue
      .Remove .count - 1
    Wend
    '����ָ��������
    For intCounter = .count To intValue - 1
      Set colTemp = .Add(intCounter)
      colTemp.Visible = True
    Next
  End With
End Property

' ***********************************************************************************************
'Ԥ������������ʽ
' ***********************************************************************************************
Private Sub SetStyles()
  Dim tdbgrd As Object
  Set tdbgrd = TrueDBGridControl
  On Error Resume Next
  With tdbgrd.Styles
    'ͨ�ñ�����ʽ����/��
    With .Item("Normal")
      .VerticalAlignment = 2 'dbgVertCenter
      .WrapText = True
    End With
    '�б�����ʽ��ǳ��/��
    With .Item("Heading")
      .Alignment = 2 'dbgCenter
      .WrapText = True
    End With
    '�н�ע��ʽ��ǳ��/��
    With .Item("Footing")
      .WrapText = True
    End With
    '���������ʽ��ǳ��/�ڣ����壬����
    With .Item("Caption")
      .Font.Name = "����_GB2312"
      .Font.Bold = True
      .Font.Size = 18
      .Alignment = 2 'dbgCenter
      .WrapText = True
    End With
    'ѡ���л�����ʽ����/��
    With .Item("Selected")
      .BackColor = vbBlue
      .ForeColor = vbWhite
      .WrapText = True
    End With
    '����������ʽ����/��
    With .Item("HighlightRow")
      .BackColor = vbBlack
      .ForeColor = vbWhite
      .WrapText = True
    End With
    '��������ʽ����/��
    With .Item("OddRow")
      .WrapText = True
    End With
    'ż������ʽ��ǳ��/��
    With .Item("EvenRow")
      .BackColor = &H80000016
      .WrapText = True
    End With
    '��������ʽ����/��
    With .Add("RedOddRow")
      .ForeColor = vbRed
      .WrapText = True
    End With
    '��������ʽ��ǳ��/��
    With .Add("RedEvenRow")
      .BackColor = &H80000016
      .ForeColor = vbRed
      .WrapText = True
    End With
    '��������ʽ����/��
    With .Add("CyanRow")
      .BackColor = vbCyan
      .WrapText = True
    End With
    '��������ʽ����/��
    With .Add("RedCyanRow")
      .BackColor = vbCyan
      .ForeColor = vbRed
      .WrapText = True
    End With
    '������ʾ��ʽ��ǳ��/��
    With .Add("Tip")
      .BackColor = vbInfoBackground
      .ForeColor = vbInfoText
      .WrapText = True
    End With
    '��ɾ���ļ�¼��ʽ����/ǳ��
    With .Add("Deleted")
      .ForeColor = vbButtonFace
      .WrapText = True
    End With
  End With
End Sub

' ***********************************************************************************************
'����������߲�������
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
  'ɾ��������
  While tdbgrd.Splits.count > 1
    tdbgrd.Splits.Remove tdbgrd.Splits.count - 1
  Wend
  With tdbgrd.Splits(0)
    .SizeMode = 0 'dbgScalable
    .Size = 1
  End With
  '���Ӳ��
  For intSplit = 1 To intSplits - 1
    Set splTemp = tdbgrd.Splits.Add(intSplit)
    splTemp.RecordSelectors = False
  Next
  '���õ�ǰ���
  tdbgrd.Split = 0
  '�����������������������У����ڽ������мӽ���������
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
'�ұ���������������������У����ڽ������мӽ���������
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
'���������
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
'�����ұ���
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
'���������ӡԤ���˵�
' ***********************************************************************************************
'Private Sub SetPrintInfo()
'  Dim tdbgrd As Object
'  Set tdbgrd = TrueDBGridControl
'  With tdbgrd.PrintInfo
'    .PreviewCaption = "��ӡԤ��"
'    .SetMenuText dbgpApsCaption, "����ҳ������"
'    .SetMenuText dbgpApsHorzCenter, "����(&C)"
'    .SetMenuText dbgpApsHorzLeft, "�����(&L)"
'    .SetMenuText dbgpApsHorzRight, "�Ҷ���(&R)"
'    .SetMenuText dbgpApsHorzText, "�ɴ�ӡ�����ˮƽλ��"
'    .SetMenuText dbgpApsMargins, "����ҳ�߾࣬ʹ�ɴ�ӡ�������(&M)"
'    .SetMenuText dbgpApsOrientation, "ѡ�����ֽ�ŷ���(&O)"
'    .SetMenuText dbgpApsPageSetup, "ҳ������(&P)"
'    .SetMenuText dbgpApsText, "ֽ�ſɴ�ӡ����̫С�����ɲ��´�ӡ���ݡ�����/������������ã�������ȷ�ϡ���ť����ҳ�����ò���С��ӡ���ݡ�"
'    .SetMenuText dbgpApsVertBottom, "�¶���(&B)"
'    .SetMenuText dbgpApsVertCenter, "����(&E)"
'    .SetMenuText dbgpApsVertText, "�ɴ�ӡ����Ĵ�ֱλ��"
'    .SetMenuText dbgpApsVertTop, "�϶���(&T)"
'    .SetMenuText dbgpDlgPagesCancel, "ȡ��(&C)"
'    .SetMenuText dbgpDlgPagesCaption, "��ӡָ��ҳ"
'    .SetMenuText dbgpDlgPagesOk, "ȷ��(&O)"
'    .SetMenuText dbgpDlgPagesPrompt, "�������ӡҳ�ŷ�Χ(-3,6,8-11,20-):"
'    .SetMenuText dbgpErrFileFormat, "�ļ���ʽ����"
'    .SetMenuText dbgpErrFileOpen, "���ļ�ʧ�ܡ�"
'    .SetMenuText dbgpErrFileSave, "�����ļ�ʧ�ܡ�"
'    .SetMenuText dbgpFilemaskAll, "�����ļ�"
'    .SetMenuText dbgpFilemaskArx, "Arx�ļ�"
'    .SetMenuText dbgpMenuActualSize, "����(&S)    Shift+Enter"
'    .SetMenuText dbgpMenuExit, "�˳�(&X)          Alt+F4"
'    .SetMenuText dbgpMenuFile, "�ļ�(&F)"
'    .SetMenuText dbgpMenuFileClose, "�ر�(&C)          Ctrl+F4"
'    .SetMenuText dbgpMenuFit, "��ҳ(&W)    Enter"
'    .SetMenuText dbgpMenuFitWidth, "ҳ��(&D)    Ctrl+Enter"
'    .SetMenuText dbgpMenuOpen, "��(&O)"
'    .SetMenuText dbgpMenuPageSetup, "ҳ������..."
'    .SetMenuText dbgpMenuPgFirst, "��ҳ(&F)    Home"
'    .SetMenuText dbgpMenuPgLast, "ĩҳ(&L)    End"
'    .SetMenuText dbgpMenuPgNext, "��ҳ(&N)    PageDown"
'    .SetMenuText dbgpMenuPgPrev, "��ҳ(&P)    PageUp"
'    .SetMenuText dbgpMenuPrint, "��ӡ(&P)          Ctrl+P"
'    .SetMenuText dbgpMenuPrintCurrPage, "��ӡ��ǰҳ(&C)    Ctrl+C"
'    .SetMenuText dbgpMenuPrintSomePages, "��ӡָ��ҳ(&S)    Ctrl+S"
'    .SetMenuText dbgpMenuSaveAs, "����Ϊ(&A)"
'    .SetMenuText dbgpMenuView, "��ͼ(&V)"
'    .SetMenuText dbgpMenuZoomIn, "�Ŵ�(&I)    +"
'    .SetMenuText dbgpMenuZoomOut, "��С(&O)    -"
'    .SetMenuText dbgpTipActualSize, "����"
'    .SetMenuText dbgpTipFit, "��ҳ"
'    .SetMenuText dbgpTipFitWidth, "ҳ��"
'    .SetMenuText dbgpTipOpen, "��"
'    .SetMenuText dbgpTipPageOf, "ҳ��"
'    .SetMenuText dbgpTipPageSetup, "ҳ������"
'    .SetMenuText dbgpTipPgFirst, "��ҳ"
'    .SetMenuText dbgpTipPgLast, "ĩҳ"
'    .SetMenuText dbgpTipPgNext, "��ҳ"
'    .SetMenuText dbgpTipPgPrev, "��ҳ"
'    .SetMenuText dbgpTipPrint, "��ӡ"
'    .SetMenuText dbgpTipSaveAs, "����Ϊ"
'    .SetMenuText dbgpTipStop, "ֹͣ����"
'    .SetMenuText dbgpTipZoom, "���ű���"
'    .SetMenuText dbgpTipZoomIn, "�Ŵ�"
'    .SetMenuText dbgpTipZoomOut, "��С"
'    .PageHeaderFont.Size = 20
'    .PageHeaderFont.Bold = True
'    .RepeatColumnHeaders = True
'    .PageFooter = "\t��\pҳ ��\Pҳ"
'    .PreviewPageOf = " ��\pҳ ��\Pҳ"
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