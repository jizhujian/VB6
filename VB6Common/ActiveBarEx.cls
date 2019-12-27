VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveBarEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mlngActiveBarObjPtr As Long

Public Sub Init(ByVal ab As Object)
  If TypeName(ab) = "ActiveBar3" Then
    mlngActiveBarObjPtr = ObjPtr(ab)
  End If
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get ActiveBarControl() As Object
  Dim objPtr2Object As New Ptr2Object
  Set ActiveBarControl = objPtr2Object.ObjectFromPtr(mlngActiveBarObjPtr)
  Set objPtr2Object = Nothing
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub SetGeneral()
  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Set ab = ActiveBarControl
  With ab
    .AlignToForm = True
    .AutoSizeChildren = 2 ' ddASClientArea
    With .Font
      .Name = "宋体"
      .Size = 9
    End With
    With .ChildBandFont
      .Name = "宋体"
      .Size = 9
    End With
    With .ControlFont
      .Name = "宋体"
      .Size = 9
    End With
    .ToolbarStyle = 2 'ddTSOffice2003
  End With
End Sub

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function GetNextToolID() As Long
  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Set ab = ActiveBarControl
  GetNextToolID = GetMaxToolID(ab) + 1
End Function

Private Function GetMaxToolID(ByVal ab As Object) As Long

  Dim objBand As Object ' ActiveBar3LibraryCtl.Band
  Dim objTool As Object ' ActiveBar3LibraryCtl.Tool
  Dim lngMaxToolID As Long

  For Each objTool In ab.Tools
    If (lngMaxToolID < objTool.ID) Then
      lngMaxToolID = objTool.ID
    End If
  Next

  For Each objBand In ab.Bands
    For Each objTool In objBand.Tools
      If (lngMaxToolID < objTool.ID) Then
        lngMaxToolID = objTool.ID
      End If
    Next
  Next

  GetMaxToolID = lngMaxToolID

End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Function GetActiveBarBandByType(ByVal intBandType As Integer) As Object   ' ActiveBar3LibraryCtl.Band
  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Dim objBand As Object ' ActiveBar3LibraryCtl.Band
  Set ab = ActiveBarControl
  For Each objBand In ab.Bands
    If objBand.Type = intBandType Then
      Set GetActiveBarBandByType = objBand
      Exit Function
    End If
  Next
End Function

' ***********************************************************************************************
' 主菜单
' ***********************************************************************************************
Public Function GetMenuBar() As Object ' ActiveBar3LibraryCtl.Band
  Set GetMenuBar = GetActiveBarBandByType(1) ' ddBTMenuBar
End Function

' ***********************************************************************************************
' 子菜单
' ***********************************************************************************************
Public Function GetPopup(ByVal PopupName As String) As Object ' ActiveBar3LibraryCtl.Band
  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Dim objBand As Object ' ActiveBar3LibraryCtl.Band
  Set ab = ActiveBarControl
  On Error Resume Next
  Set objBand = ab.Bands(PopupName)
  On Error GoTo 0
  If (Not objBand Is Nothing) Then
    If objBand.Type = 2 Then ' ddBTPopup
      Set GetPopup = objBand
    End If
  End If
End Function

' ***********************************************************************************************
' 工具栏
' ***********************************************************************************************
Public Function GetToolbar() As Object ' ActiveBar3LibraryCtl.Band
  Set GetToolbar = GetActiveBarBandByType(0) ' ddBTNormal
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Function AddActiveBarBand(ByVal strBandName As String, ByVal intBandType As Integer, Optional ByVal strBandCaption As String) As Object   ' ActiveBar3LibraryCtl.Band
  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Dim objBand As Object ' ActiveBar3LibraryCtl.Band
  Set ab = ActiveBarControl
  Set objBand = ab.Bands.Add(strBandName)
  With objBand
    .Type = intBandType
    .Caption = strBandCaption
    .GrabHandleStyle = 0 'ddGSNone
  End With
  Set AddActiveBarBand = objBand
End Function

' ***********************************************************************************************
' 新增主菜单
' ***********************************************************************************************
Public Function AddMenuBar(ByVal MenuBarName As String, Optional ByVal MenuBarCaption As String) As Object ' ActiveBar3LibraryCtl.Band
  Set AddMenuBar = AddActiveBarBand(MenuBarName, 1, MenuBarCaption)
End Function

' ***********************************************************************************************
' 新增子菜单
' ***********************************************************************************************
Public Function AddPopup(ByVal PopupName As String, Optional ByVal PopupCaption As String) As Object ' ActiveBar3LibraryCtl.Band
  Set AddPopup = AddActiveBarBand(PopupName, 2, PopupCaption)
End Function

' ***********************************************************************************************
' 新增工具栏
' ***********************************************************************************************
Public Function AddToolbar(ByVal ToolbarName As String, Optional ByVal ToolbarCaption As String) As Object ' ActiveBar3LibraryCtl.Band
  Dim objBand As Object ' ActiveBar3LibraryCtl.Band
  Set objBand = AddActiveBarBand(ToolbarName, 0, ToolbarCaption)
  With objBand
    .flags = 1 + 2 + 4 + 8 + 256 + 512 'ddBFDockTop + ddBFDockBottom + ddBFDockLeft + ddBFDockRight + ddBFFixed + ddBFStretch
    .DisplayMoreToolsButton = False
    .WrapTools = True
  End With
  Set AddToolbar = objBand
End Function

' ***********************************************************************************************
' 新增分隔符
' ***********************************************************************************************
Public Function AddSeparator(ByVal Band As Object, Optional ByVal index As Long = -1) As Object  ' ActiveBar3LibraryCtl.Tool

  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Dim objTool As Object ' ActiveBar3LibraryCtl.Tool

  Set ab = ActiveBarControl
  On Error Resume Next
  Set objTool = ab.Tools("JZJBandToolSeparator")
  On Error GoTo 0
  If (objTool Is Nothing) Then
    Set objTool = ab.Tools.Add(GetNextToolID, "JZJBandToolSeparator")
    With objTool
      .ControlType = 5 'ddTTSeparator
      .Category = "Separator"
      .Caption = "Separator"
    End With
  End If

  With Band.Tools
    .Insert IIf(index = -1, .count, index), objTool
  End With

  Set AddSeparator = objTool

End Function

Public Function CreateTool(ByVal ToolName As String, ByVal ToolCaption As String, _
  Optional ByVal ToolCategory As String, Optional ByVal ToolIcon As String, Optional ByVal CaptionPosition As Integer = 4) As Object ' ActiveBar3LibraryCtl.Tool

  Dim ab As Object ' ActiveBar3LibraryCtl.ActiveBar3
  Dim objTool As Object ' ActiveBar3LibraryCtl.Tool
  Dim objIconResource As IconResource

  Set ab = ActiveBarControl
  On Error Resume Next
  Set objTool = ab.Tools(ToolName)
  On Error GoTo 0
  If (objTool Is Nothing) Then
    Set objTool = ab.Tools.Add(GetNextToolID, ToolName)
    With objTool
      .ControlType = 0 'ddTTButton
      .CaptionPosition = CaptionPosition
      .Style = 3 ' ddSIconText
      .Caption = ToolCaption
      If ToolCaption > "" Then
        .Category = ToolCategory
      End If
      If (ToolIcon > "") Then
        Set objIconResource = New IconResource
        .SetPicture 0, objIconResource.LoadResIcon(ToolIcon)
        Set objIconResource = Nothing
        .ImageWidth = 240
        .ImageHeight = 240
      End If
    End With
  End If
  Set CreateTool = objTool

End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub AddBandTool(ByVal Band As Object, ByVal Tool As Object, Optional ByVal index As Long = -1)
  With Band.Tools
    .Insert IIf(index = -1, .count, index), Tool
  End With
End Sub

Public Function IndexOfBandTool(ByVal Band As Object, ByVal ToolName As String) As Long
  Dim toolIndex As Long
  With Band.Tools
    For toolIndex = 0 To .count - 1
      If UCase(.Item(toolIndex).Name) = UCase(ToolName) Then
        IndexOfBandTool = toolIndex
        Exit Function
      End If
    Next
  End With
  IndexOfBandTool = -1
End Function

Public Sub BandOpen(ByVal Band As Object, ByVal Cancel As Object)
  If Band.Name = "SysCustomize" Then
    Cancel.Value = True
  End If
End Sub
