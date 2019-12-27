VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "actbar3.ocx"
Begin VB.Form frmMain 
   Caption         =   "图标资源浏览器"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   7905
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin ActiveBar3LibraryCtl.ActiveBar3 ab 
      Align           =   1  'Align Top
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7908
      _LayoutVersion  =   2
      _ExtentX        =   13944
      _ExtentY        =   9737
      _DataPath       =   ""
      Bands           =   "frmMain.frx":038A
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iconKeys() As String
Private pages As Integer
Private currentPage As Integer
Private Const CountPerPage = 50

Private Sub ab_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
If Band.Name = "SysCustomize" Then
    'Stops system customize menu from being displayed
    Cancel = True
End If
End Sub

Private Sub ab_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)
  Dim i As Integer
  Dim page As Integer
  If Left(Tool.Name, 6) = "Pages_" And IsNumeric(Mid(Tool.Name, 7)) Then
    page = CInt(Mid(Tool.Name, 7))
    If page = currentPage Then Exit Sub
    If currentPage > 0 Then
      ab.Bands("Pages").Tools("Pages_" & currentPage).Checked = False
    End If
    With ab.Bands("Toolbar").Tools
      .RemoveAll
      For i = (page - 1) * CountPerPage To page * CountPerPage - 1
        .Insert .Count, ab.Tools("Seperator")
        If i > UBound(iconKeys) Then Exit For
        .Insert .Count, ab.Tools(iconKeys(i))
      Next
    End With
    ab.Bands("Pages").Tools("Pages_" & page).Checked = True
    currentPage = page
    ab.RecalcLayout
  Else
    Clipboard.Clear
    Clipboard.SetText Tool.Caption
  End If
End Sub

Private Sub Form_Load()

  Dim file As New dotNET2COM.file
  iconKeys = file.ReadAllLines(App.Path & "\IconKeyList.txt")
  Set file = Nothing

  Dim i As Integer
  With ab.Tools
    .Add(999, "Seperator").ControlType = ddTTSeparator
    For i = LBound(iconKeys) To UBound(iconKeys)
      With .Add(1000 + i, iconKeys(i))
        .Caption = iconKeys(i)
        .Style = ddSIconText
        .CaptionPosition = ddCPBelow
        .SetPicture ddITNormal, LoadResIcon(iconKeys(i))
      End With
    Next
  End With

  Dim math As New dotNET2COM.math
  pages = math.Ceiling((UBound(iconKeys) + 1) / CountPerPage)

  With ab.Bands("Pages").Tools
    For i = 1 To pages
      .Insert .Count, ab.Tools("Seperator")
      With .Add(100 + i * 2 + 1, "Pages_" & i)
        .Caption = "第" & i & "页"
        .Style = ddSText
      End With
    Next
  End With

  ab_ToolClick ab.Bands("Pages").Tools("Pages_1")

End Sub
