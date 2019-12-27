VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TPropertySheet 
   BackStyle       =   0  '透明
   ClientHeight    =   2568
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4572
   ScaleHeight     =   2568
   ScaleWidth      =   4572
   ToolboxBitmap   =   "PropertySheetC.ctx":0000
   Begin MSComctlLib.ImageList StdImages 
      Left            =   600
      Top             =   2160
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0312
            Key             =   "frame"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0424
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0495
            Key             =   "dots"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":058F
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0689
            Key             =   "check_off"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":079B
            Key             =   "check_on"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":08AD
            Key             =   "minus"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstCheck 
      Appearance      =   0  'Flat
      Height          =   24
      ItemData        =   "PropertySheetC.ctx":091B
      Left            =   480
      List            =   "PropertySheetC.ctx":0922
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   2112
      Width           =   972
   End
   Begin VB.TextBox txtList 
      Appearance      =   0  'Flat
      Height          =   492
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1560
      Width           =   732
   End
   Begin MSComCtl2.MonthView monthView 
      Height          =   2088
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   3984
      _ExtentX        =   7027
      _ExtentY        =   3683
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483635
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   72155137
      TitleBackColor  =   -2147483635
      TitleForeColor  =   -2147483634
      CurrentDate     =   36675
   End
   Begin MSComCtl2.UpDown UpDown 
      Height          =   372
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   216
      _ExtentX        =   360
      _ExtentY        =   656
      _Version        =   393216
      OrigLeft        =   240
      OrigTop         =   1560
      OrigRight       =   456
      OrigBottom      =   1932
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdBrowse 
      Height          =   372
      Left            =   240
      Picture         =   "PropertySheetC.ctx":0929
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.ListBox lstBox 
      Appearance      =   0  'Flat
      Height          =   204
      Left            =   1080
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.TextBox txtBox 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   492
   End
   Begin MSComctlLib.ImageList StdImages2 
      Left            =   240
      Top             =   1560
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0A13
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0A81
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0AF2
            Key             =   "dots"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0BEC
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0CE6
            Key             =   "check_on_sel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0DF8
            Key             =   "check_off_sel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":0F0A
            Key             =   "check_on_2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":101C
            Key             =   "check_off_2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":112E
            Key             =   "check_off"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":1240
            Key             =   "check_on"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetC.ctx":1352
            Key             =   "frame"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fGrid 
      Height          =   1956
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   3450
      _Version        =   393216
      GridColor       =   12632256
      FocusRect       =   2
      GridLinesUnpopulated=   1
      ScrollBars      =   2
      BorderStyle     =   0
      Appearance      =   0
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "TPropertySheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************************************************************
'
'枕善居汉化收藏整理
'发布日期：05/07/05
'描  述：组件属性窗口控件 Ver1.0
'网  站：http://www.codesky.net/
'
'
'****************************************************************************
' *******************************************************
' 控件         : TPropertySheet.Ctl
' 作者         : Marclei V Silva (MVS)
' 程序员       : Marclei V Silva (MVS) [Spnorte Consultoria de Inform醫ica]
' 编写日期     : 06/16/2000 -- 09:08:30
' 描   述      : PropertySheet control which show property
'              : and values in a spredsheet like
'              : workspace
' *******************************************************
Option Explicit
Option Compare Text

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Top As Long
    Left As Long
    Right As Long
    Bottom As Long
End Type

Private Type CELL_RECT
    Top As Long
    Left As Long
    Width As Long
    Height As Long
    ButtonLeft As Long
    ButtonTop As Long
    ButtonWidth As Long
    ButtonHeight As Long
    WindowLeft As Long
    WindowTop As Long
    WindowWidth As Long
    InterfaceLeft As Long
End Type

Private rc As CELL_RECT

' Keep up with the errors
Const g_ErrConstant As Long = vbObjectError + 1000
Const m_constClassName = "PropertySheet.TPropertySheet"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

Const ColCount = 4
Private Const flexSortStringNoCaseAsending = 5
Private Const flexSortNumericAscending = 3

' indicates the columns of th grid
Private Enum enPropertyColumns
    ColStatus = 0
    colName = 1
    colValue = 2
    colSort = 3
End Enum

Private Const COL_WIDTH = 18
Private Const ROW_HEIGHT = 18

' 默认属性值:
Const m_def_Appearance = 1
Const m_def_BorderStyle = 1
Const m_def_ForeColor = &H80000008
Const m_def_GridColor = &HC0C0C0
Const m_def_BackColor = &HFFFFFF
Const m_def_SelBackColor = &H8000000D
Const m_def_SelForeColor = &H8000000E
Const m_def_CatBackColor = &HFFC080
Const m_def_CatForeColor = &HFFFFFF
Const m_def_ShowToolTips = 0
Const m_def_Enabled = 1
'Const m_def_AllowEmptyValues = 1
Const m_def_ExpandableCategories = 1
Const m_def_NameWidth = 0
Const m_def_RequiresEnter = 0
Const m_def_ShowCategories = 1
Const m_def_ItemHeight = 8
Const m_def_ExpandedImage = 0
Const m_def_CollapsedImage = 0
Const m_def_Initializing = 0

' 属性变量:
Private m_ExpandedImage As Integer
Private m_CollapsedImage As Integer
Private m_CatFont As Font
Private m_ForeColor As OLE_COLOR
Private m_GridColor As OLE_COLOR
Private m_BackColor As OLE_COLOR
Private m_SelBackColor As OLE_COLOR
Private m_SelForeColor As OLE_COLOR
Private m_CatBackColor As OLE_COLOR
Private m_CatForeColor As OLE_COLOR
Private m_ShowToolTips As Boolean
Private m_SelectedItem As Object
Private m_Enabled As Boolean
Private m_Font As Font
'Private m_AllowEmptyValues As Boolean
Private m_ExpandableCategories As Boolean
Private m_NameWidth As Single
Private m_RequiresEnter As Boolean
Private m_ShowCategories As Boolean

' 私有声明
Private m_bEditFlag As Boolean
Private m_EditRow As Integer
Private m_BrowseWnd As Object
Private m_bDataChanged As Boolean
Private m_bBrowseMode As Boolean
Private m_OldValue As Variant
Private m_bDirty As Boolean
Private m_strBuffer As String
Private m_Categories As TCategories
Private m_LastKey As Integer
Private m_ItemHeight As Integer
Private m_SelectedRow As Integer
Private m_lPadding As Single
Private m_Properties As Collection
Private m_strText As String
Private m_bUserMode As Boolean
Private m_hIml As Long
Private m_hImlStd As Long
Private m_lIconSize As Long
Private m_bListDirty As Boolean

' 事件声明:
Event Browse(ByVal Left, ByVal Top, ByVal Width, ByVal Prop As TProperty)
Event CategoryCollapsed(Cancel As Boolean)
Event CategoryExpanded(Cancel As Boolean)
Event EnterEditMode(ByVal Prop As TProperty, Cancel As Boolean)
Attribute EnterEditMode.VB_Description = "Occurs when the edit control is to be shown allowing the user to edit the property"
Event GetDisplayString(ByVal Prop As TProperty, DisplayString As String, UseDefault As Boolean)
Attribute GetDisplayString.VB_Description = "Occurs when the control needs the display string of a property. This event is called only if the property has the FormatProperty set to ""CustomeDisplay"""
Event ParseString(ByVal Prop As TProperty, ByVal Text As String, UseDefault As Boolean)
Attribute ParseString.VB_Description = "Occurs when the user changes a property that has the format property set to ""CustomDisplay"""
Event BeforePropertyChanged(ByVal Prop As TProperty, NewValue, Cancel As Boolean)
Event AfterPropertyChanged(ByVal Prop As TProperty, NewValue)
Event SelectionChanged(ByVal Prop As TProperty)
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event EditError(ErrMessage As String)
Event HideControls()
Event BrowseForFile(ByVal Prop As TProperty, Title As String, ByRef InitDir As String, Filter As String, FilterIndex As Integer, flags As Long)
Event BrowseForFolder(ByVal Prop As TProperty, Title As String, Path As String, Prompt As String)
Event OnClear()

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Appearance
Public Property Get Appearance() As psAppearanceSettings
    Appearance = UserControl.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As psAppearanceSettings)
    UserControl.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As psBorderStyle
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As psBorderStyle)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Properties(Index As Variant) As TProperty
    Dim Ptr As Long
    ' get a pointer to the property
    Ptr = m_Properties(Index)
    ' retrieve property object
    Set Properties = ObjectFromPtr(Ptr)
End Property

Public Property Get PropertyCount() As Long
    PropertyCount = m_Properties.Count
End Property

Public Property Get Parent() As Object
    Set Parent = Extender.Parent
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
'Public Property Get AllowEmptyValues() As Boolean
'    AllowEmptyValues = m_AllowEmptyValues
'End Property
'
'Public Property Let AllowEmptyValues(ByVal New_AllowEmptyValues As Boolean)
'    m_AllowEmptyValues = New_AllowEmptyValues
'    PropertyChanged "AllowEmptyValues"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ExpandableCategories() As Boolean
    ExpandableCategories = m_ExpandableCategories
End Property

Public Property Let ExpandableCategories(ByVal New_ExpandableCategories As Boolean)
    m_ExpandableCategories = New_ExpandableCategories
    PropertyChanged "ExpandableCategories"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get NameWidth() As Single
    NameWidth = m_NameWidth
End Property

Public Property Let NameWidth(ByVal New_NameWidth As Single)
    m_NameWidth = New_NameWidth
    PropertyChanged "NameWidth"
    Grid_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get RequiresEnter() As Boolean
    RequiresEnter = m_RequiresEnter
End Property

Public Property Let RequiresEnter(ByVal New_RequiresEnter As Boolean)
    m_RequiresEnter = New_RequiresEnter
    PropertyChanged "RequiresEnter"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowCategories() As Boolean
    ShowCategories = m_ShowCategories
End Property

Public Property Let ShowCategories(ByVal New_ShowCategories As Boolean)
    m_ShowCategories = New_ShowCategories
    PropertyChanged "ShowCategories"
    Grid_ShowCategories
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ShowToolTips() As Boolean
    ShowToolTips = m_ShowToolTips
End Property

Public Property Let ShowToolTips(ByVal New_ShowToolTips As Boolean)
    m_ShowToolTips = New_ShowToolTips
    PropertyChanged "ShowToolTips"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFC080&
Public Property Get CatBackColor() As OLE_COLOR
    CatBackColor = m_CatBackColor
End Property

Public Property Let CatBackColor(ByVal New_CatBackColor As OLE_COLOR)
    Dim i As Integer
    Dim Row As Integer
    
    m_CatBackColor = New_CatBackColor
    PropertyChanged "CatBackColor"
    With fGrid
        .Redraw = False
        For i = 1 To m_Categories.Count
            Row = m_Categories(i).Row
            .Row = Row
            .Col = colName
            .ColSel = colValue
            .CellBackColor = New_CatBackColor
        Next
        .Redraw = True
    End With
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get CatForeColor() As OLE_COLOR
    CatForeColor = m_CatForeColor
End Property

Public Property Let CatForeColor(ByVal New_CatForeColor As OLE_COLOR)
    Dim i As Integer
    Dim Row As Integer
    
    m_CatForeColor = New_CatForeColor
    PropertyChanged "CatForeColor"
    With fGrid
        .Redraw = False
        For i = 1 To m_Categories.Count
            Row = m_Categories(i).Row
            .Row = Row
            .Col = colName
            .ColSel = colValue
            .CellForeColor = New_CatForeColor
        Next
        .Redraw = True
    End With
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000D&
Public Property Get SelBackColor() As OLE_COLOR
    SelBackColor = m_SelBackColor
End Property

Public Property Let SelBackColor(ByVal New_SelBackColor As OLE_COLOR)
    m_SelBackColor = New_SelBackColor
    PropertyChanged "SelBackColor"
    Hilite m_SelectedRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000E&
Public Property Get SelForeColor() As OLE_COLOR
    SelForeColor = m_SelForeColor
End Property

Public Property Let SelForeColor(ByVal New_SelForeColor As OLE_COLOR)
    m_SelForeColor = New_SelForeColor
    PropertyChanged "SelForeColor"
    Hilite m_SelectedRow
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get Categories() As TCategories
    Set Categories = m_Categories
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=9,0,0,0
Public Property Get SelectedItem() As Object
    Set SelectedItem = m_SelectedItem
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFC080&
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Integer
    Dim j As Integer
    Dim Row As Integer
    
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
    With fGrid
        .BackColor = New_BackColor
        .BackColorBkg = New_BackColor
        .BackColorUnpopulated = New_BackColor
        .GridColorFixed = New_BackColor
        .BackColorFixed = New_BackColor
    End With
    With fGrid
        .Redraw = False
        For i = 1 To m_Categories.Count
            For j = 1 To m_Categories(i).Properties.Count
                Row = m_Categories(i).Properties(j).Row
                .Row = Row
                .Col = colName
                .ColSel = colValue
                .CellBackColor = New_BackColor
            Next
        Next
        .Redraw = True
    End With
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Dim i As Integer
    Dim j As Integer
    
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    With fGrid
        .Redraw = False
        For i = 1 To m_Categories.Count
            For j = 1 To m_Categories(i).Properties.Count
                .Row = m_Categories(i).Properties(j).Row
                .Col = colName
                .ColSel = colValue
                .CellForeColor = New_ForeColor
            Next
        Next
        .Redraw = True
    End With
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&HFF
Public Property Get GridColor() As OLE_COLOR
    GridColor = m_GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    m_GridColor = New_GridColor
    PropertyChanged "GridColor"
    fGrid.GridColor = New_GridColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get CatFont() As Font
    Set CatFont = m_CatFont
End Property

Public Property Set CatFont(ByVal New_CatFont As Font)
    Dim i As Integer
    Dim Row As Integer
    
    Set m_CatFont = New_CatFont
    PropertyChanged "CatFont"
    With fGrid
        .Redraw = False
        For i = 1 To m_Categories.Count
            Row = m_Categories(i).Row
            .Row = Row
            .Col = colName
            .ColSel = colValue
            .CellFontName = New_CatFont.Name
            .CellFontBold = New_CatFont.Bold
            .CellFontItalic = New_CatFont.Italic
            .CellFontStrikeThrough = New_CatFont.Strikethrough
            .CellFontUnderline = New_CatFont.Underline
            .CellFontSize = New_CatFont.Size
        Next
        .Redraw = True
    End With
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ExpandedImage() As Integer
    ExpandedImage = m_ExpandedImage
End Property

Public Property Let ExpandedImage(ByVal New_ExpandedImage As Integer)
    m_ExpandedImage = New_ExpandedImage
    PropertyChanged "ExpandedImage"
    'Grid_ChangeCategoryImage New_ExpandedImage, True
    Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ItemHeight() As Integer
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Integer)
    m_ItemHeight = New_ItemHeight
    PropertyChanged "ItemHeight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get CollapsedImage() As Integer
    CollapsedImage = m_CollapsedImage
End Property

Public Property Let CollapsedImage(ByVal New_CollapsedImage As Integer)
    m_CollapsedImage = New_CollapsedImage
    PropertyChanged "CollapsedImage"
    'Grid_ChangeCategoryImage New_CollapsedImage, False
    Grid_Paint
End Property

Public Property Let ImageList(ByRef vImageList As Variant)
    m_hIml = 0
    If (VarType(vImageList) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (Err.Number = 0) Then
            ' OK
            m_hIml = PtrFromObject(vImageList)
        Else
            Debug.Print "Failed to Get Image list Handle", "PropertySheet.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        Dim cx As Long, cy As Long
        If (ImageList_GetIconSize(vImageList.hImageList, cx, cy) <> 0) Then
            m_lIconSize = cy
        End If
    End If
'    PropertyChanged "ImageList"
'    Grid_Paint
End Property

Private Function Image_List(hIml As Long) As MSComctlLib.ImageList
    Set Image_List = ObjectFromPtr(hIml)
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    m_bDirty = True
    Grid_Paint
End Sub

Private Sub lstCheck_ItemCheck(Item As Integer)
    If m_bListDirty = True Then Exit Sub
    UpdateCheckList
End Sub

Private Sub monthView_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            UpdateProperty monthView.Value, True
            KeyCode = 0
        Case vbKeyEscape
            monthView.Visible = False
            m_bBrowseMode = False
            KeyCode = 0
    End Select
End Sub

Private Sub txtBox_GotFocus()
    If IsProperty(m_SelectedItem) Then
        If m_SelectedItem.ValueType <> psTime Then
            SelectText
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    ' recalc padding
    RecalcPadding
    Set m_CatFont = Ambient.Font
    m_Enabled = m_def_Enabled
'    m_AllowEmptyValues = m_def_AllowEmptyValues
    m_ExpandableCategories = m_def_ExpandableCategories
    m_NameWidth = m_def_NameWidth
    m_RequiresEnter = m_def_RequiresEnter
    m_ShowCategories = m_def_ShowCategories
    m_ShowToolTips = m_def_ShowToolTips
    m_CatBackColor = m_def_CatBackColor
    m_CatForeColor = m_def_CatForeColor
    m_SelBackColor = m_def_SelBackColor
    m_SelForeColor = m_def_SelForeColor
    m_ItemHeight = m_def_ItemHeight
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_GridColor = m_def_GridColor
    m_ExpandedImage = m_def_ExpandedImage
    m_CollapsedImage = m_def_CollapsedImage
'    m_ImageList = m_def_ImageList
    m_bDirty = True
    UserControl.BorderStyle = m_def_BorderStyle
    UserControl.Appearance = m_def_Appearance
    fGrid.Clear
    fGrid.Rows = 0
    fGrid.cols = ColCount
    m_Categories.Clear
    Grid_Config     ' config the grid
    Grid_Resize     ' resize the grid
    fGrid.ColWidth(ColStatus) = COL_WIDTH * Screen.TwipsPerPixelX
    If (UserControl.Ambient.UserMode = False) Then
        m_bUserMode = False
        Set m_Properties = Nothing
        Set m_Properties = New Collection
        With m_Categories.Add("TPropertySheet", "组件属性窗口控件")
            .Properties.Add "Name", "名称", UserControl.Ambient.DisplayName
            With .Properties.Add("Selected", "项", "值")
                .Selected = True
            End With
        End With
    Else
        m_bUserMode = True
    End If
End Sub

Private Sub UserControl_LostFocus()
    HideControls
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' get default Font
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    ' recalc padding
    RecalcPadding
    Set m_CatFont = PropBag.ReadProperty("CatFont", Ambient.Font)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
'    m_AllowEmptyValues = PropBag.ReadProperty("AllowEmptyValues", m_def_AllowEmptyValues)
    m_ExpandableCategories = PropBag.ReadProperty("ExpandableCategories", m_def_ExpandableCategories)
    m_NameWidth = PropBag.ReadProperty("NameWidth", m_def_NameWidth)
    m_RequiresEnter = PropBag.ReadProperty("RequiresEnter", m_def_RequiresEnter)
    m_ShowCategories = PropBag.ReadProperty("ShowCategories", m_def_ShowCategories)
    m_ShowToolTips = PropBag.ReadProperty("ShowToolTips", m_def_ShowToolTips)
    m_CatBackColor = PropBag.ReadProperty("CatBackColor", m_def_CatBackColor)
    m_CatForeColor = PropBag.ReadProperty("CatForeColor", m_def_CatForeColor)
    m_SelBackColor = PropBag.ReadProperty("SelBackColor", m_def_SelBackColor)
    m_SelForeColor = PropBag.ReadProperty("SelForeColor", m_def_SelForeColor)
    m_ItemHeight = PropBag.ReadProperty("ItemHeight", m_def_ItemHeight)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_GridColor = PropBag.ReadProperty("GridColor", m_def_GridColor)
    m_ExpandedImage = PropBag.ReadProperty("ExpandedImage", m_def_ExpandedImage)
    m_CollapsedImage = PropBag.ReadProperty("CollapsedImage", m_def_CollapsedImage)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    fGrid.Clear
    fGrid.Rows = 0
    fGrid.cols = ColCount
    m_Categories.Clear
    Grid_Config     ' config the grid
    Grid_Resize     ' resize the grid
    If m_ShowCategories = True Then
        fGrid.ColWidth(ColStatus) = COL_WIDTH * Screen.TwipsPerPixelX
    Else
        fGrid.ColWidth(ColStatus) = 0
    End If
    If (UserControl.Ambient.UserMode = False) Then
        m_bUserMode = False
        Set m_Properties = Nothing
        Set m_Properties = New Collection
        With m_Categories.Add("TPropertySheet", "组件属性窗口控件")
            .Properties.Add "Name", "名称", UserControl.Ambient.DisplayName
            With .Properties.Add("Selected", "项", "值")
                .Selected = True
            End With
        End With
    Else
        m_bUserMode = True
    End If
End Sub

Private Sub UserControl_Show()
    Static s_bNotFirst As Boolean
    Dim hwndParent As Long
    If Not (s_bNotFirst) Then
        ' set the parent of this resources
' 支持嵌套控件 季祝建 2008-04-21
        'SetParent lstCheck.hwnd, Extender.Parent.hwnd
'        SetParent lstBox.hwnd, Extender.Parent.hwnd
'        SetParent monthView.hwnd, Extender.Parent.hwnd
'        SetParent txtList.hwnd, Extender.Parent.hwnd
        'SetParent lstCheck.hwnd, Extender.Parent.hwnd
        hwndParent = GetParent(hwnd)
        'SetParent lstCheck.hwnd, hwndParent
        SetParent lstBox.hwnd, hwndParent
        SetParent monthView.hwnd, hwndParent
        SetParent txtList.hwnd, hwndParent
        ' stay on top
        StayOnTop lstBox.hwnd
        StayOnTop monthView.hwnd
        StayOnTop txtList.hwnd
        s_bNotFirst = True
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("AllowEmptyValues", m_AllowEmptyValues, m_def_AllowEmptyValues)
    Call PropBag.WriteProperty("ExpandableCategories", m_ExpandableCategories, m_def_ExpandableCategories)
    Call PropBag.WriteProperty("NameWidth", m_NameWidth, m_def_NameWidth)
    Call PropBag.WriteProperty("RequiresEnter", m_RequiresEnter, m_def_RequiresEnter)
    Call PropBag.WriteProperty("ShowCategories", m_ShowCategories, m_def_ShowCategories)
    Call PropBag.WriteProperty("ShowToolTips", m_ShowToolTips, m_def_ShowToolTips)
    Call PropBag.WriteProperty("CatBackColor", m_CatBackColor, m_def_CatBackColor)
    Call PropBag.WriteProperty("CatForeColor", m_CatForeColor, m_def_CatForeColor)
    Call PropBag.WriteProperty("SelBackColor", m_SelBackColor, m_def_SelBackColor)
    Call PropBag.WriteProperty("SelForeColor", m_SelForeColor, m_def_SelForeColor)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("GridColor", m_GridColor, m_def_GridColor)
    Call PropBag.WriteProperty("CatFont", m_CatFont, Ambient.Font)
    Call PropBag.WriteProperty("ExpandedImage", m_ExpandedImage, m_def_ExpandedImage)
    Call PropBag.WriteProperty("CollapsedImage", m_CollapsedImage, m_def_CollapsedImage)
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, m_def_ItemHeight)
End Sub

Private Sub UserControl_Initialize()
On Error Resume Next
    m_bDirty = True
    m_NameWidth = 0
    m_hIml = 0
    ' initialize objects
    Set m_Categories = New TCategories
    ' initialize the object
    m_Categories.Init Me
    ' create properties collection
    Set m_Properties = New Collection
    ' initialize grid
    Grid_Initialize
    m_hImlStd = PtrFromObject(StdImages)
End Sub

Private Sub Grid_Initialize()
On Error Resume Next
    ' set grid parameters for the sheet
    With fGrid
        .Redraw = False
        .Left = 0
        .Top = 0
        .FixedRows = 0
        .cols = ColCount
        .FixedCols = 1
        .ColWidth(colSort) = 0
        .GridLines = flexGridFlat
        .GridLinesFixed = flexGridNone
        .SelectionMode = flexSelectionByRow
        .FillStyle = flexFillRepeat
        .FocusRect = flexFocusNone
        .GridLines = flexGridFlat
        .Font.Name = "Verdana"
        .MergeCells = flexMergeFree
        .MergeCol(ColStatus) = True
        .Redraw = True
    End With
End Sub

Private Sub UserControl_Paint()
    On Error GoTo Err_UserControl_Paint
    On Error Resume Next
    If m_bDirty = False Then Exit Sub
    Grid_Paint
    m_bDirty = False
    
    Exit Sub
Err_UserControl_Paint:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UserControl_Paint")
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Err_UserControl_Resize
On Error Resume Next
    Grid_Resize

    Exit Sub
Err_UserControl_Resize:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UserControl_Resize")
End Sub

Private Sub UserControl_Terminate()
    Set m_Categories = Nothing
    Set m_Properties = Nothing
End Sub

Private Sub Grid_Config()
    On Error GoTo Err_Grid_Config

    With fGrid
        .Redraw = False
        .GridColor = m_GridColor
        .GridLines = flexGridFlat
        .GridColorFixed = m_BackColor
        .FixedCols = 1
        .GridColorUnpopulated = vbBlue
        .BackColorFixed = m_BackColor
        .BackColorSel = m_SelBackColor
        .BackColorBkg = m_BackColor
        .BackColorUnpopulated = m_BackColor
        .BackColor = m_BackColor
        .ForeColorSel = m_SelForeColor
        .ForeColor = m_ForeColor
        Set .Font = m_Font
        .Redraw = True
    End With

    Exit Sub
Err_Grid_Config:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_Config")
End Sub

Private Sub Grid_Resize()
    On Error GoTo Err_Grid_Resize

    Dim wid As Single  ' column width - col(3) width
    Dim gw As Single   ' grid line width
    Dim ch As Single   ' cell height
    Dim sb As Single   ' spacer (scrollbar)
    Dim Cell As Integer
    Dim nw As Single
    Dim cols As Single
    Dim h As Double
    Dim edge As Single
    
    ' hide all the controls visible
    HideControls
    ' update grid columns
    With fGrid
        ' avoid flickering
        .Redraw = False
        ' save current cell
        Cell = Cell_Save
        ' update grid rect
        .Left = 0
        .Top = 0
        fGrid.Width = UserControl.ScaleWidth
        fGrid.Height = UserControl.ScaleHeight
        ' get grid line width in screen resolution
        gw = .GridLineWidth * Screen.TwipsPerPixelX
        ' get the cell height
        ch = .CellHeight
        On Error Resume Next
        wid = 0
        cols = 0
        ' check for ShowCategories property
        If m_ShowCategories Then
            .ColWidth(ColStatus) = COL_WIDTH * Screen.TwipsPerPixelX
            wid = wid + .ColWidth(ColStatus)
            cols = cols + 1
        Else
            .ColWidth(ColStatus) = 0
        End If
        ' detect name width here
        If m_NameWidth = 0 Then
            nw = Cell_NameWidth '+ (colName * gw)
        Else
            nw = m_NameWidth
        End If
        ' update the name column width
        .ColWidth(colName) = nw
        ' get column 2 width
        wid = wid + .ColWidth(colName)
        ' increase columns
        cols = cols + 1
        ' sort column is invisible
        .ColWidth(colSort) = 0
        If ScrollBarVisible(.hwnd) Then
            ' If the contents don't fit in the available outline space,
            ' then we have to compensate for the width of the scrollbar.
            sb = Screen.TwipsPerPixelX * (GetSystemMetrics(SM_CXVSCROLL) + GetSystemMetrics(SM_CXBORDER))
        Else
            ' Otherwise, we don't have a scrollbar.
            sb = 0
        End If
        ' set value column width now
        .ColWidth(colValue) = (fGrid.Width - (wid + sb + ((cols + 1) * gw))) + ((cols - 1) * gw) + 2 * Screen.TwipsPerPixelX
        ' restore cell position
        Cell_Restore Cell
        ' start drawing from here
        StoreCellPosition
        .Redraw = True
    End With

    Exit Sub
Err_Grid_Resize:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_Resize")
End Sub

Private Sub Grid_Clear()
    On Error GoTo Err_Grid_Clear

    With fGrid
        .Clear
        .Rows = 0
        .cols = ColCount
    End With

    Exit Sub
Err_Grid_Clear:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_Clear")
End Sub

Private Sub Grid_Edit(KeyAscii As Integer, bFocus As Boolean)
    On Error GoTo Err_Grid_Edit

    Dim Cancel As Boolean
    If TypeName(m_SelectedItem) <> "TProperty" Then Exit Sub
    If m_SelectedItem.ReadOnly = True Then Exit Sub
    Cancel = False
    RaiseEvent EnterEditMode(m_SelectedItem, Cancel)
    If Cancel = True Then Exit Sub
    ' set last key to nothing
    m_LastKey = 0
    ' give way to windows (good after a event raise)
    DoEvents
    m_EditRow = fGrid.Row
    m_bEditFlag = True
    If IsObject(m_SelectedItem.Value) Then
        Set m_OldValue = m_SelectedItem.Value
    Else
        m_OldValue = m_SelectedItem.Value
    End If
    fGrid.Col = colValue
    ShowTextBox
    If Not IsWindowLess(m_SelectedItem) Then
        UpDown.Visible = False
        ShowBrowseButton
    Else
        cmdBrowse.Visible = False
        If IsIncremental(m_SelectedItem) Then
            ShowUpDown
        Else
            UpDown.Visible = False
        End If
    End If
    m_bDataChanged = False
    If bFocus = True Then
        If txtBox.Visible = False Then ShowTextBox
        txtBox.SetFocus
    End If
    Select Case KeyAscii
        Case 0 To Asc(" ")
            txtBox.SelStart = 0
            txtBox.SelLength = Len(txtBox.Text)
        Case Else
            txtBox.Text = Chr(KeyAscii)
            txtBox.SelStart = 1
    End Select
    Exit Sub
Err_Grid_Edit:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_Edit")
End Sub

Friend Sub AddNewCategory(ByVal objCat As TCategory)
     On Error GoTo Err_AddNewCategory

    Dim CurrRow As Integer
    Dim Ptr As Long
    Dim Index As Long
    Dim Cell As Integer
    
    ' stop flickering
    fGrid.Redraw = False
    ' save row col position
    Cell = Cell_Save
    ' hide controls
    HideControls
    ' dehilite
'    DeHilite
    ' add new category
    fGrid.AddItem "  " & vbTab & "" & vbTab & "" & vbTab & "" '& vbTab & ""
    ' set the row to update
    CurrRow = fGrid.Rows - 1
    objCat.Row = CurrRow
    ' get the pointer to the object
    Ptr = objCat.Handle
    ' get the catego
    Index = MakeDWord(objCat.Index, 0)
    ' row data will contain object pointer
    fGrid.RowData(CurrRow) = Ptr
    ' this row will be merged
    fGrid.MergeRow(CurrRow) = True
    Row_Category objCat
    'fGrid.Row = CurrRow
    'fGrid.Col = colSort
    'fGrid.Text = Index
    Grid_Index
    ' restore cell position
    Cell_Restore Cell
    ' active drawing
    fGrid.Redraw = True
    ' give way to windows
    DoEvents
    
    Exit Sub
Err_AddNewCategory:
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "AddNewCategory")
End Sub

Private Sub Row_Category(objCat As TCategory)
    Dim Row As Integer
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR
    Dim ObjFont As StdFont
    
    Row = objCat.Row
    Set ObjFont = m_CatFont
    With fGrid
        .Row = Row
        .Col = ColStatus
        .CellPictureAlignment = flexAlignCenterCenter
        If objCat.Expanded = False Then
            Set .CellPicture = StdImages.ListImages("plus").Picture
        Else
            Set .CellPicture = StdImages.ListImages("minus").Picture
        End If
        If m_ShowCategories = True Then
            .RowHeight(Row) = DefaultHeight
        Else
            .RowHeight(Row) = 0
        End If
        ' configure back color
        GetObjectColors objCat, tmpBackColor, tmpForeColor
        If m_hIml <> 0 Then
' 显示分类图像BUG 季祝建 2008-04-21
            ' set the current state for this category
            ' expanded/collapseed
'            If objCat.Expanded Then
'                objCat.Image = m_ExpandedImage
'            Else
'                objCat.Image = m_CollapsedImage
'            End If
            Cell_DrawPicture Row, colName, objCat.Image
        End If
        .Col = colName
        .ColSel = colValue
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForeColor
        .CellAlignment = flexAlignLeftCenter
        If Not ObjFont Is Nothing Then
            .CellFontName = ObjFont.Name
            .CellFontBold = ObjFont.Bold
            .CellFontItalic = ObjFont.Italic
            .CellFontStrikeThrough = ObjFont.Strikethrough
            .CellFontUnderline = ObjFont.Underline
            .CellFontSize = ObjFont.Size
        End If
        .Text = Pad(objCat.Caption)
    End With
End Sub

Private Sub GetObjectColors(obj As Object, Back_Color As OLE_COLOR, Optional Fore_Color As OLE_COLOR)
    If obj.Selected Then
        Back_Color = m_SelBackColor
        Fore_Color = m_SelForeColor
    Else
        If IsProperty(obj) Then
            Back_Color = m_BackColor
            Fore_Color = m_ForeColor
        Else
            Back_Color = m_CatBackColor
            Fore_Color = m_CatForeColor
        End If
        If obj.BackColor <> CLR_INVALID Then
            Back_Color = obj.BackColor
        End If
        If obj.ForeColor <> CLR_INVALID Then
            Fore_Color = obj.ForeColor
        End If
    End If
End Sub

Private Sub AddNewProperty(ByVal objProp As TProperty, ByVal Relative As TCategory)
    On Error GoTo Err_AddNewProperty

    Dim strText As String
    Dim CurrRow As Integer
    Dim Index As Long
    Dim Ptr As Long
    Dim Cell As Integer
    
    ' stop flickering
    fGrid.Redraw = False
    ' save row col position
    Cell = Cell_Save
    ' hide controls
    HideControls
    ' dehilite
'    DeHilite
    ' update grid
    With fGrid
        ' add a new item
        .AddItem "" & vbTab & "" & vbTab & "" & vbTab & 0 '"" & vbTab & "" & vbTab & 0
        ' get property handle
        Ptr = objProp.Handle
        ' create unique index for this property
        Index = MakeDWord(Relative.Index, objProp.Index)
        ' save curr row
        CurrRow = .Rows - 1
        ' save row to the property object
        objProp.Row = CurrRow
        ' set handle to rowData to retrive the object later
        .RowData(CurrRow) = Ptr
        Row_Property objProp, Relative.Expanded
        ' sort column
        '.Row = CurrRow
        '.Col = colSort
        '.Text = Index
        Grid_Index
    End With
    ' add object pointer to properties collection
    m_Properties.Add Ptr, objProp.Caption
    ' restore cell position
    Cell_Restore Cell
    ' active drawing
    fGrid.Redraw = True
    ' give way to windows
    DoEvents

    Exit Sub
Err_AddNewProperty:
    ' active drawing
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "AddNewProperty")
End Sub

Private Sub Row_Property(objProp As TProperty, bExpanded As Boolean)
    Dim strDisplayText As String
    Dim Cell As Integer
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR
    Dim Row As Integer
    Dim strText As String
    
    Row = objProp.Row
    ' get property name caption
    strText = Pad(objProp.Caption)
    ' get display Text
    strDisplayText = GetDisplayString(objProp)
    With fGrid
        .Row = Row
        .Col = colName
        ' if ShowCategories is enabled check row height
        If m_ShowCategories = True Then
            ' if category is not expanded the cell height must be zero
            If bExpanded Then
                .RowHeight(Row) = DefaultHeight
            Else
                .RowHeight(Row) = 0    ' invisible
            End If
        Else
            .RowHeight(Row) = DefaultHeight
        End If
        ' configure back color
        GetObjectColors objProp, tmpBackColor, tmpForeColor
        If m_hIml <> 0 Then
            ' draw picture inside column #1
            If .RowHeight(Row) <> 0 Then
                Cell_DrawPicture Row, colName, objProp.Image
            End If
        End If
        If objProp.BackColor <> CLR_INVALID Then
            .CellBackColor = tmpBackColor
        End If
        If objProp.ForeColor <> CLR_INVALID Then
            .CellForeColor = tmpForeColor
        End If
        .CellAlignment = flexAlignLeftCenter
        .Text = strText
        ' value text
        If HasGraphicInterface(objProp) Then
            DrawGraphicInterface objProp, Row
            strDisplayText = Space(m_lPadding) & strDisplayText
        End If
        .Col = colValue
        If objProp.BackColor <> CLR_INVALID Then
            .CellBackColor = tmpBackColor
        End If
        If objProp.ForeColor <> CLR_INVALID Then
            .CellForeColor = tmpForeColor
        End If
        .CellAlignment = flexAlignLeftCenter
        .Text = strDisplayText
    End With
End Sub

Private Sub fGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If TypeOf m_SelectedItem Is TCategory Then
        Select Case KeyCode
            Case vbKeyLeft
                If m_SelectedItem.Expanded = True Then
                    CollapseCategory m_SelectedItem
                End If
            Case vbKeyRight
                If m_SelectedItem.Expanded = False Then
                    ExpandCategory m_SelectedItem
                End If
            Case Else
                txtBox_KeyDown KeyCode, Shift
        End Select
    End If
End Sub

Private Sub fGrid_LostFocus()
    If m_bDataChanged Then
        ' update value only if we don't have a browse
        ' window active
        If m_bBrowseMode = False Then
            UpdateProperty txtBox.Text
        End If
    End If
End Sub

Private Sub fGrid_Scroll()
    HideControls
End Sub

Private Sub fGrid_RowColChange()
    On Error GoTo Err_fGrid_RowColChange

    Dim CurrRow As Integer
    
    ' get current row
    CurrRow = fGrid.Row
    ' check row changed here
    If m_bDataChanged Then
        If m_RequiresEnter = True Then
            UpdateProperty m_OldValue
        Else
            If IsProperty(m_SelectedItem) Then
                If (m_bBrowseMode = False) Then
                    If (m_SelectedItem.ListValues.Count = 0) Then
                        UpdateProperty txtBox.Text
                    Else
                        On Error Resume Next
                        If m_SelectedItem.ValueType = psCombo Then
                            UpdateProperty txtBox.Text
                        Else
                            If m_SelectedItem.ListValues.Exists(txtBox.Text) Then
                                UpdateProperty m_SelectedItem.ListValues(txtBox.Text).Value
                            End If
                        End If
                    End If
                End If
            End If
        End If
        m_bDataChanged = False
    End If
    ' highlight current row no matter the type
    '    HideControls
    '    HideBrowseWnd
    Hilite CurrRow
    If IsProperty(m_SelectedItem) Then
        ' hide all controls
        HideBrowseWnd
        ' over a property then fire this event
        RaiseEvent SelectionChanged(m_SelectedItem)
        fGrid_KeyPress Asc(" ")
    Else
        HideControls
    End If
    Exit Sub

Err_fGrid_RowColChange:
End Sub

Private Sub fGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Err_fGrid_MouseMove

    ' if show tips is disabled then exit
    If m_ShowToolTips = False Or m_Categories.Count = 0 Then Exit Sub
    
    Static LastRow As Integer
    Dim Row As Integer
    
    ' get the row at mouse position
    Row = fGrid.MouseRow
    ' if this is the last row then exit
    If LastRow = Row Then Exit Sub
    ' save this row
    LastRow = Row
    Dim objTemp As Object
    ' get the appropriate tip from row's object
    Set objTemp = GetRowObject(Row)
    ' sets the tip
    If Not objTemp Is Nothing Then
        fGrid.TooltipText = objTemp.TooltipText
    End If

    Exit Sub
Err_fGrid_MouseMove:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "fGrid_MouseMove")
End Sub

Private Sub fGrid_DblClick()
    On Error GoTo Err_fGrid_DblClick

    Dim Row As Integer
    
    If m_bBrowseMode Or m_Categories.Count = 0 Then Exit Sub
    ' get mouse row
    Row = fGrid.MouseRow
    ' if we are in a property cell
    If IsProperty(m_SelectedItem) Then
        If m_SelectedItem.ReadOnly = False Then
            ' check if we have to browse
            If IsBrowsable(m_SelectedItem) Then
                ' browse the property
                BrowseProperty
            Else
                ' get next avail row
                GetNextVisibleRowValue
            End If
        Else
            ' hide the text box
            If txtBox.Visible Then
                txtBox.Visible = False
            End If
        End If
        RaiseEvent DblClick
    Else
        ' toggle category state expanded/collapsed
        ToggleCategoryState
    End If
    If txtBox.Visible Then
        txtBox.SetFocus
    End If
    Exit Sub

Err_fGrid_DblClick:
End Sub

Private Sub fGrid_Click()
    On Error GoTo Err_fGrid_Click
    
    If m_bBrowseMode Or m_Categories.Count = 0 Then Exit Sub
    
    Dim Col As Integer
    
    ' get mouse coordinates in row/col
    Col = fGrid.MouseCol
    ' if its a category and the column is 0 then
    ' promote collapse/expand
    If m_SelectedItem Is Nothing Then
        Hilite fGrid.MouseRow
    End If
    If Not IsProperty(m_SelectedItem) Then
        If Col = 0 Then
            ToggleCategoryState
        End If
    Else
        If m_SelectedItem.ReadOnly = True Then
            HideControls
        End If
    End If

    Exit Sub
Err_fGrid_Click:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "fGrid_Click")
End Sub

Private Sub fGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo Err_txtBox_KeyPress
    
    Dim FindString As String
    Dim i As Integer
    Dim n As Integer
    
    If m_Categories.Count = 0 Then Exit Sub
' 在分类上按键BUG 季祝建 2008-04-21
    If Not IsProperty(m_SelectedItem) Then Exit Sub
    ' update last key value
    m_LastKey = KeyAscii
    If (m_SelectedItem.ValueType = psBoolean Or _
       m_SelectedItem.ValueType = psDropDownList) And _
       KeyAscii > 32 Then
        FindString = Chr(KeyAscii)
        n = Len(FindString)
        For i = 1 To m_SelectedItem.ListValues.Count
            If Left(m_SelectedItem.ListValues(i).Caption, n) = FindString Then
                Exit For
            End If
        Next
        If i <= m_SelectedItem.ListValues.Count Then
            txtBox.Text = m_SelectedItem.ListValues(i).Caption
            txtBox.SetFocus
            txtBox.SelStart = 0
            txtBox.SelLength = Len(m_SelectedItem.ListValues(i).Caption)
        End If
        KeyAscii = 0
    Else
        Grid_Edit KeyAscii, True
    End If

    Exit Sub
Err_txtBox_KeyPress:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtBox_KeyPress")
End Sub

Private Sub txtBox_Click()
    If m_SelectedItem.ValueType = psDropDownList Or m_SelectedItem.ValueType = psLongText Or m_SelectedItem.ValueType = psDropDownCheckList Then
        If m_bBrowseMode = False Then
            BrowseProperty
        End If
    End If
End Sub

Private Sub txtBox_KeyPress(KeyAscii As Integer)
    On Error GoTo Err_txtBox_KeyPress

    Dim FindString As String
    Dim i As Integer
    Dim n As Integer
    
    ' update last key
    m_LastKey = KeyAscii
    
    If (m_SelectedItem.ValueType = psBoolean Or _
       m_SelectedItem.ValueType = psDropDownList) Then
        FindString = Chr(KeyAscii)
        n = Len(FindString)
        For i = 1 To m_SelectedItem.ListValues.Count
            If Left(m_SelectedItem.ListValues(i).Caption, n) = FindString Then
                Exit For
            End If
        Next
        If i <= m_SelectedItem.ListValues.Count Then
            txtBox.Text = m_SelectedItem.ListValues(i).Caption
            txtBox.SetFocus
            txtBox.SelStart = 0
            txtBox.SelLength = Len(m_SelectedItem.ListValues(i).Caption)
        End If
        KeyAscii = 0
    Else
        If KeyAscii = vbKeyReturn Then KeyAscii = 0
    End If

    Exit Sub
Err_txtBox_KeyPress:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtBox_KeyPress")
End Sub

Private Sub txtBox_LostFocus()
    If m_bDataChanged Then
        ' update value only if we don't have a browse
        ' window active
        If m_bBrowseMode = False Then
            UpdateProperty txtBox.Text
        End If
    End If
End Sub

Private Sub txtBox_DblClick()
    On Error GoTo Err_txtBox_DblClick

    GetNextVisibleRowValue

    Exit Sub
Err_txtBox_DblClick:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtBox_DblClick")
End Sub

Private Sub txtBox_Change()
    On Error GoTo Err_txtBox_Change

    ' text has changed
    m_bDataChanged = True
    
    Exit Sub
Err_txtBox_Change:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtBox_Change")
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_txtBox_KeyDown

    Dim NewRow As Integer
    Dim AltDown As Boolean
    
    ' alt key was pressed?
    AltDown = (Shift And vbAltMask) > 0
    Select Case KeyCode
        Case vbKeyDelete
            ' if its an object or image clean it up
            If m_SelectedItem.ValueType = psObject Or _
               m_SelectedItem.ValueType = psPicture Then
                UpdateProperty Nothing, True
            ElseIf m_SelectedItem.ValueType = psLongText Then
                UpdateProperty "", True
            End If
        
        Case vbKeyEscape
            m_bDataChanged = False
            KeyCode = 0
            fGrid.SetFocus
        
        Case vbKeyReturn
            If IsWindowLess(m_SelectedItem) Or m_SelectedItem.ValueType = psCombo Then
                UpdateProperty txtBox.Text
            End If
            fGrid_RowColChange
        
        Case vbKeyDown
            If AltDown Then
                ' check if we have to browse here
                If m_SelectedItem.ValueType = psDropDownList Or _
                   m_SelectedItem.ValueType = psBoolean Or _
                   m_SelectedItem.ValueType = psDropDownCheckList Or _
                   m_SelectedItem.ValueType = psLongText Or _
                   IsArray(m_SelectedItem.Value) Or _
                   m_SelectedItem.ValueType = psCombo Or m_SelectedItem.ValueType = psDate Then
                    BrowseProperty
                ElseIf IsIncremental(m_SelectedItem) Then
                    UpdateUpDown -m_SelectedItem.UpDownIncrement
                End If
            Else
                If (m_RequiresEnter = True And m_bDataChanged = False) Or m_RequiresEnter = False Then
                    NewRow = GetNextVisibleRow
                    If NewRow <> -1 Then
                        fGrid.SetFocus
                        fGrid.Row = NewRow
                        fGrid.Refresh
                        fGrid_RowColChange
                        'fGrid_Click
                    End If
                Else
                    Beep
                    KeyCode = 0
                End If
            End If
        Case vbKeyUp
            If AltDown Then
                If IsIncremental(m_SelectedItem) Then
                    UpdateUpDown m_SelectedItem.UpDownIncrement
                End If
            Else
                If (m_RequiresEnter = True And m_bDataChanged = False) Or m_RequiresEnter = False Then
                    NewRow = GetPreviousVisibleRow
                    If NewRow <> -1 Then
                        fGrid.SetFocus
                        fGrid.Row = NewRow
                        fGrid.Refresh
                        fGrid_RowColChange
                        'fGrid_Click
                    End If
                Else
                    Beep
                    KeyCode = 0
                End If
            End If
    End Select

    Exit Sub
Err_txtBox_KeyDown:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtBox_KeyDown")
End Sub

Private Sub txtList_Change()
    m_bDataChanged = True
End Sub

Private Sub txtList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub txtList_LostFocus()
    If m_bDataChanged Then
        Dim tmpValue As String
        tmpValue = StripBkLinefeed(txtList.Text)
        If IsArray(m_SelectedItem.Value) Then
            UpdateProperty Split(tmpValue, vbCrLf)
        Else
            UpdateProperty tmpValue
        End If
    End If
End Sub

Private Sub txtList_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_txtList_KeyDown

    Dim CtrlDown As Boolean
    
    ' alt key was pressed?
    CtrlDown = (Shift And vbCtrlMask) > 0
    Select Case KeyCode
        Case vbKeyEscape
            m_bDataChanged = False
            fGrid_RowColChange
        Case vbKeyReturn
            If CtrlDown = False Then
                txtList_LostFocus
                fGrid_RowColChange
            End If
    End Select

    Exit Sub
Err_txtList_KeyDown:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "txtList_KeyDown")
End Sub

Private Sub lstBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_lstBox_KeyDown

    Dim Index As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            ' get list box index to access listvalues value
            Index = lstBox.ListIndex + 1
            If Index > 0 Then
                UpdateProperty m_SelectedItem.ListValues(Index).Value
            End If
        Case vbKeyEscape
            lstBox.Visible = False
            m_bBrowseMode = False
        Case vbKeyUp
            Index = lstBox.ListIndex - 1
            If Index >= 0 Then
                txtBox.Text = lstBox.List(Index)
            End If
        Case vbKeyDown
            Index = lstBox.ListIndex + 1
            If Index <= lstBox.ListCount Then
                txtBox.Text = lstBox.List(Index)
            End If
    End Select

    Exit Sub
Err_lstBox_KeyDown:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "lstBox_KeyDown")
End Sub

Private Sub lstBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Err_lstBox_MouseUp
    
    Dim Index As Integer
    
    ' get list box index to access listvalues value
    Index = lstBox.ListIndex + 1
    If Index > 0 Then
        UpdateProperty m_SelectedItem.ListValues(Index).Value, True
    End If
    
    Exit Sub
Err_lstBox_MouseUp:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "lstBox_MouseUp")
End Sub

Private Sub lstCheck_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_lstCheck_KeyDown

    Dim Index As Integer
    
    Select Case KeyCode
        Case 32
            UpdateCheckList
        Case vbKeyReturn
            fGrid.SetFocus
        Case vbKeyEscape
            lstCheck.Visible = False
            m_bBrowseMode = False
    End Select

    Exit Sub
Err_lstCheck_KeyDown:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "lstCheck_KeyDown")
End Sub

Private Sub UpdateCheckList()
    On Error GoTo Err_UpdateCheckList
    
    Dim Value As String
    Dim i As Integer
    
    For i = 0 To lstCheck.ListCount - 1
        If lstCheck.Selected(i) Then
            If Value = "" Then
                Value = m_SelectedItem.ListValues(i + 1).Value
            Else
                Value = Value & Chr(0) & m_SelectedItem.ListValues(i + 1).Value
            End If
        End If
    Next
    UpdateProperty Value, True
    
    Exit Sub
Err_UpdateCheckList:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UpdateCheckList")
End Sub

Private Function GetNextVisibleRowValue()
    On Error GoTo Err_GetNextVisibleRowValue

    Dim strCaption As String
    Dim i As Integer
    Dim n As Integer
    
    ' exit if there's no item value
    If m_SelectedItem.ListValues.Count = 0 Then Exit Function
    n = 1
    ' loop the item values to find the next candidate
    For i = 1 To m_SelectedItem.ListValues.Count
        strCaption = txtBox.Text
        If m_SelectedItem.ListValues(i).Caption = strCaption Then
            n = i + 1
            If n > m_SelectedItem.ListValues.Count Then
                n = 1
            End If
            Exit For
        End If
    Next
    ' update property
    UpdateProperty m_SelectedItem.ListValues(n).Value, True
    Exit Function

Err_GetNextVisibleRowValue:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "GetNextVisibleRowValue")
End Function

Private Sub UpDown_DownClick()
    On Error GoTo Err_UpDown_DownClick
    
    ' update with decreasing increment value
    UpdateUpDown -m_SelectedItem.UpDownIncrement
    
    Exit Sub
Err_UpDown_DownClick:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UpDown_DownClick")
End Sub

Private Sub UpDown_UpClick()
    On Error GoTo Err_UpDown_UpClick
    
    ' update with increasing increment
    UpdateUpDown m_SelectedItem.UpDownIncrement
    
    Exit Sub
Err_UpDown_UpClick:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UpDown_UpClick")
End Sub

Private Sub UpdateUpDown(Increment As Variant)
    Dim Value As Variant
    Dim StartPos As Integer
    
    Value = m_SelectedItem.Value
    If m_SelectedItem.ValueType = psTime Then
        Dim Interval As String
        If txtBox.SelStart >= 0 And txtBox.SelStart <= 2 Then
            StartPos = 0
            Interval = "h"
        ElseIf txtBox.SelStart >= 3 And txtBox.SelStart <= 5 Then
            StartPos = 3
            Interval = "n"
        Else
            StartPos = 6
            Interval = "s"
        End If
        Value = DateAdd(Interval, Increment, Value)
    Else
        Dim MinValue As Variant
        Dim MaxValue As Variant
        m_SelectedItem.GetRange MinValue, MaxValue
        Value = Value + Increment
'        If Not IsEmpty(MinValue) Then
'            If Value < MinValue Then
'                Value = MinValue
'            End If
'        End If
'        If Not IsEmpty(MaxValue) Then
'            If Value > MaxValue Then
'                Value = MaxValue
'            End If
'        End If
    End If
    UpdateProperty Value, True
    If m_SelectedItem.ValueType = psTime Then
        On Error Resume Next
        txtBox.SetFocus
        txtBox.SelStart = StartPos
        txtBox.SelLength = 2
    Else
        SelectText
    End If
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo Err_cmdBrowse_Click

    BrowseProperty

    Exit Sub
Err_cmdBrowse_Click:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "cmdBrowse_Click")
End Sub

Private Sub HideControls()
    Dim i As Integer
    
    On Error Resume Next
    RaiseEvent HideControls
    For i = 0 To UserControl.Controls.Count - 1
        If UserControl.Controls(i).Name <> "fGrid" Then
            UserControl.Controls(i).Visible = False
        End If
    Next
    m_bEditFlag = False
    m_bBrowseMode = False
End Sub

Public Sub UpdateProperty(ByVal NewValue As Variant, Optional bForceUpdate As Boolean = False)
    On Error GoTo Err_UpdateProperty
    
    ' If there is nothing to update then exit
    If (m_SelectedItem Is Nothing) Or _
       (m_bDataChanged = False And bForceUpdate = False) Then Exit Sub
    
    ' hide the controls activated by Grid_Edit()
    '    If m_SelectedItem.ValueType <> psDropDownCheckList Then
    '        'HideControls
    '    End If
    ' check for AllowEmptyValues
    If m_SelectedItem.AllowEmptyValues = False And IsVarEmpty(NewValue) Then
        RaiseEvent EditError("“" & m_SelectedItem.Caption & "”值不能为空。")
        Exit Sub
    End If
    ' check requires enter here
    '    If m_RequiresEnter = True And m_LastKey <> vbKeyReturn Then
    '        ' back to previous row
    '        RaiseEvent EditError("Value for '" & m_SelectedItem.Caption & "' can only be updated with the ENTER key press")
    '        Exit Sub
    '    End If
    Dim Cancel As Boolean
    
    ' cancel is false
    Cancel = False
    ' get permission to change the property
    RaiseEvent BeforePropertyChanged(m_SelectedItem, NewValue, Cancel)
    ' permission denied get out
    If Cancel = True Then
        fGrid.SetFocus
        Exit Sub
    End If
    'StopFlicker hwnd
    fGrid.Redraw = False
    Dim tmpValue As Variant
    ' data changed is false now
    m_bDataChanged = False
    ' check for a passed object here
    If IsObject(NewValue) Then
        Set m_SelectedItem.Value = NewValue
        RaiseEvent AfterPropertyChanged(m_SelectedItem, NewValue)
    Else
        tmpValue = ConvertValue(NewValue, m_SelectedItem.ValueType)
        If Not IsNull(tmpValue) Then
            If IsIncremental(m_SelectedItem) Then
                Dim MinValue
                Dim MaxValue
                m_SelectedItem.GetRange MinValue, MaxValue
                If Not IsEmpty(MinValue) Or Not IsEmpty(MaxValue) Then
                    If Not IsEmpty(MinValue) Then
                        If tmpValue < MinValue Then
                            tmpValue = MinValue
                        End If
                    End If
                    If Not IsEmpty(MaxValue) Then
                        If tmpValue > MaxValue Then
                            tmpValue = MaxValue
                        End If
                    End If
                End If
            End If
            m_SelectedItem.Value = tmpValue
            RaiseEvent AfterPropertyChanged(m_SelectedItem, tmpValue)
        Else
            RaiseEvent EditError("不能更新“" & m_SelectedItem.Caption & "”，无效的数据类型。")
        End If
    End If
    ' update textbox text value
    If m_SelectedItem.ValueType <> psDropDownCheckList Then
        UpdateTextBox m_SelectedItem
    End If
    fGrid.Redraw = True
    
    Exit Sub
Err_UpdateProperty:
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UpdateProperty")
End Sub

Private Sub BrowseProperty()
    On Error GoTo Err_BrowseProperty

    ' Are we in browse mode? (window is open)
    If m_bBrowseMode Then
        If IsObject(m_BrowseWnd) Then
            On Error Resume Next
            m_BrowseWnd.Visible = False
        End If
        m_bBrowseMode = False
    Else
        ' Now we are in browse mode
        ' a specific window is open for editting purposes
        m_bBrowseMode = True
        ' raise event indicating we are browsing now
        RaiseEvent Browse(rc.WindowLeft, rc.WindowTop, rc.WindowWidth, m_SelectedItem)
        UpdateProperty m_SelectedItem.Value, True
        ' give way to windows
        DoEvents
        ' select the edit method based on ValueType property
        Select Case m_SelectedItem.ValueType
            Case psCombo, psDropDownList, psBoolean: EditCombo
            Case psDropDownCheckList: EditCheckList
            Case psDate: EditDate
            Case psPicture: EditPicture
            Case psFont: EditFont
            Case psFile: EditFile
            Case psFolder: EditFolder
            Case psLongText: EditLongText
            Case psColor: EditColor
            Case psCustom
                If IsArray(m_SelectedItem.Value) Then
                    EditLongText
                End If
        End Select
        If txtBox.Visible = False Then
            ShowTextBox
        End If
        '        txtBox.SetFocus
    End If

    Exit Sub
Err_BrowseProperty:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "BrowseProperty")
End Sub

Private Sub EditLongText()
    On Error GoTo Err_EditLongText

    Dim strBuffer As String
        
    SetControlFont txtList
    txtList.ZOrder
    If IsArray(m_SelectedItem.Value) Then
        strBuffer = Join(m_SelectedItem.Value, vbCrLf)
    Else
        strBuffer = m_SelectedItem.Value
    End If
    txtList.Left = rc.WindowLeft
    'txtList.Top = RC.WindowTop
    txtList.Width = rc.WindowWidth
    txtList.Height = m_ItemHeight * TextHeight("A")
    txtList.Top = FixTopPos(txtList.Height)
    txtList.Text = strBuffer
    m_bDataChanged = False
    Set m_BrowseWnd = txtList
    txtList.Visible = True
    txtList.SetFocus

    Exit Sub
Err_EditLongText:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "EditLongText")
End Sub

Private Sub EditCombo()
    On Error GoTo Err_EditCombo

    Dim i As Integer
    Dim h As Integer
    
    lstBox.Clear
    lstBox.ZOrder
    SetControlFont lstBox
    For i = 1 To m_SelectedItem.ListValues.Count
        lstBox.AddItem m_SelectedItem.ListValues(i).Caption
        If m_SelectedItem.ListValues(i).Value = m_SelectedItem.Value Then
            lstBox.ListIndex = lstBox.NewIndex
        End If
    Next
    If m_SelectedItem.ListValues.Count > m_ItemHeight Then
        h = m_ItemHeight * TextHeight("A")
    Else
        h = (m_SelectedItem.ListValues.Count + 1) * TextHeight("A")
    End If
    lstBox.Left = rc.WindowLeft
    lstBox.Width = rc.WindowWidth
    lstBox.Height = h
    lstBox.Top = FixTopPos(lstBox.Height)
    Set m_BrowseWnd = lstBox
    lstBox.Visible = True
    lstBox.SetFocus
    m_bDataChanged = False
    
    Exit Sub
Err_EditCombo:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "EditCombo")
End Sub

Private Sub EditCheckList()
    On Error GoTo Err_EditCheckList

    Dim vArray As Variant
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Value As String
    Dim t As Single
    Dim Header As Single
    
    m_bListDirty = True
    lstCheck.Clear
    lstCheck.ZOrder
    SetControlFont lstCheck
    Value = m_SelectedItem.Value
    vArray = Split(Value, Chr(0))
    For i = 1 To m_SelectedItem.ListValues.Count
        lstCheck.AddItem m_SelectedItem.ListValues(i).Caption
        If Trim(m_SelectedItem.Value) <> "" Then
            If Not IsNull(vArray) Then
                For j = LBound(vArray) To UBound(vArray)
                    If StrComp(m_SelectedItem.ListValues(i).Value, vArray(j), vbTextCompare) = 0 Then
                        lstCheck.Selected(lstCheck.NewIndex) = True
                    End If
                Next
            End If
        End If
    Next
    lstCheck.ListIndex = -1
    lstCheck.Width = rc.WindowWidth
    If m_SelectedItem.ListValues.Count > m_ItemHeight Then
        h = m_ItemHeight * TextHeight("A")
    Else
        h = (m_SelectedItem.ListValues.Count + 1) * TextHeight("A")
    End If
    lstCheck.Height = h
    t = FixTopPos(lstCheck.Height)
    ' list with check box has a header so we have to skip this
    ' height off the position so that it will fit right on the
    ' screen
    Header = ((GetSystemMetrics(SM_CYCAPTION) + (GetSystemMetrics(SM_CYBORDER) * 3)) * Screen.TwipsPerPixelY)
    ' remove header off-set from the top property
    t = t - Header
    lstCheck.Top = t
    lstCheck.Left = rc.WindowLeft
    If BorderStyle = psBorderSingle Then
        lstCheck.Left = fGrid.Left + (fGrid.Width - lstCheck.Width) 'lstCheck.Left - (2 * Screen.TwipsPerPixelX)
        lstCheck.Top = lstCheck.Top - ((2 * Screen.TwipsPerPixelY))
    End If
    Set m_BrowseWnd = lstCheck
    lstCheck.Visible = True
    lstCheck.SetFocus
    m_bDataChanged = False
    m_bListDirty = False
    
    Exit Sub
Err_EditCheckList:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "EditCheckList")
End Sub

Private Sub UpdateTextBox(objProp As TProperty)
    On Error GoTo Err_UpdateTextBox

    Dim strDisplayStr As String
'    StopFlicker hwnd
    ' get the display string for cell
    strDisplayStr = GetDisplayString(objProp)
    txtBox.Text = strDisplayStr
    m_bDataChanged = False
    Cell_ValueChanged objProp, m_EditRow
    Grid_Edit 32, False
'    SelectText
'    Release
    Exit Sub
    
Err_UpdateTextBox:
'    Release
    Err.Raise Err.Number, GenErrSource(m_constClassName, "UpdateTextBox")
End Sub

Private Sub EditFont()
    On Error GoTo Err_EditFont

    Dim ObjFont As StdFont
    Dim dlgCMD As New cCommonDialog
    
    ' check if Font was specified
    If TypeOf m_SelectedItem.Value Is StdFont Then
        Set ObjFont = m_SelectedItem.Value
    Else
        ' create a new Font
        Set ObjFont = New StdFont
    End If
    With dlgCMD
        .DialogTitle = "Font"
        .FontName = ObjFont.Name
        .FontBold = ObjFont.Bold
        .FontItalic = ObjFont.Italic
        .FontSize = ObjFont.Size
        .FontStrikethru = ObjFont.Strikethrough
        .FontUnderline = ObjFont.Underline
        .flags = CF_BOTH + CF_EFFECTS
        .ShowFont
        ObjFont.Name = .FontName
        ObjFont.Size = .FontSize
        ObjFont.Bold = .FontBold
        ObjFont.Italic = .FontItalic
        ObjFont.Underline = .FontUnderline
        ObjFont.Strikethrough = .FontStrikethru
    End With
    ' update property
    UpdateProperty ObjFont, True
    ' clean up
    Set dlgCMD = Nothing
    Exit Sub
    
Err_EditFont:
    Set dlgCMD = Nothing
    m_bBrowseMode = False
End Sub

Private Sub EditPicture()
    On Error GoTo Err_EditPicture
    
    Dim strFileName As String
    Dim sTitle As String
    Dim sFilter As String
    Dim iFilterIndex As Integer
    Dim lFlags As Long
    Dim dlgCMD As New cCommonDialog
    Dim Pict As StdPicture
    Dim InitDir As String
    
    ' check if file already exists
    On Error Resume Next
    ' get filename
    strFileName = m_SelectedItem.Value
    ' check if file exist
    Dir strFileName
    ' update file name
    If Err.Number <> 0 Then
        If Len(strFileName) = 3 Then
            strFileName = strFileName & "*.*"
        Else
            strFileName = ""
        End If
    End If
    ' these vars will be passed to the user define
    sTitle = "Open Picture"
    sFilter = "Picture Files|*.bmp;*.gif;*.jpg;*.jpeg;*.wmf;*.ico;.png|All Files (*.*)|*.*"
    iFilterIndex = 1
    lFlags = OFN_FILEMUSTEXIST
    InitDir = CurDir
    ' call the event for user definition
    RaiseEvent BrowseForFile(m_SelectedItem, sTitle, InitDir, sFilter, iFilterIndex, lFlags)
    With dlgCMD
        .InitDir = InitDir
        .Filter = sFilter
        .DialogTitle = sTitle
        .FilterIndex = iFilterIndex
        .Filename = strFileName
        .flags = lFlags
        .hwnd = hwnd
        .ShowOpen
        If Len(.Filename) > 0 Then
            On Error Resume Next
            Set Pict = LoadPicture(.Filename)
            If Err = 0 Then
                UpdateProperty Pict, True
            End If
        End If
    End With
    Set dlgCMD = Nothing
    m_bBrowseMode = False

    Exit Sub
Err_EditPicture:
    Set dlgCMD = Nothing
    m_bBrowseMode = False
End Sub

Private Sub EditDate()
    On Error GoTo Err_EditDate

    Dim tmpValue As Date
    
    If IsDate(m_SelectedItem.Value) Then
        tmpValue = CDate(m_SelectedItem.Value)
    Else
        tmpValue = Date
    End If
    With monthView
        .Left = rc.WindowLeft
        .Top = rc.WindowTop
        .Top = FixTopPos(.Height)
        .Day = Day(tmpValue)
        .Month = Month(tmpValue)
        .Year = Year(tmpValue)
        .ZOrder
        .Visible = True
    End With
    monthView.SetFocus
    Set m_BrowseWnd = monthView

    Exit Sub
Err_EditDate:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "EditDate")
End Sub

Private Sub monthView_DateClick(ByVal DateClicked As Date)
    UpdateProperty DateClicked, True
End Sub

Private Sub EditFile()
    On Error GoTo Err_EditFile

    Dim strFileName As String
    Dim sTitle As String
    Dim sFilter As String
    Dim iFilterIndex As Integer
    Dim lFlags As Long
    Dim dlgCMD As New cCommonDialog
    Dim InitDir As String
    
    ' check if file already exists
    On Error Resume Next
    ' get filename
    strFileName = m_SelectedItem.Value
    ' check if file exist
    Dir strFileName
    ' update file name
    If Err.Number <> 0 Then
        If Len(strFileName) = 3 Then
            strFileName = strFileName & "*.*"
        Else
            strFileName = ""
        End If
    End If
    ' these vars will be passed to the user define
    sTitle = "Open"
    sFilter = "All files (*.*)|*.*"
    iFilterIndex = 1
    lFlags = OFN_FILEMUSTEXIST
    InitDir = CurDir
    ' call the event for user defined vars
    RaiseEvent BrowseForFile( _
       m_SelectedItem, _
       sTitle, _
       InitDir, _
       sFilter, _
       iFilterIndex, _
       lFlags)
    ' update dialog properties
    With dlgCMD
        .InitDir = InitDir
        .Filter = sFilter
        .DialogTitle = sTitle
        .FilterIndex = iFilterIndex
        .Filename = strFileName
        .CancelError = True
        .flags = lFlags
        .hwnd = hwnd
        .ShowOpen
        strFileName = .Filename
    End With
    If Len(strFileName) > 0 Then
        ' update property
        UpdateProperty strFileName, True
    End If
    ' clean dialog object
    Set dlgCMD = Nothing
    
    Exit Sub
Err_EditFile:
    Set dlgCMD = Nothing
    m_bBrowseMode = False
End Sub

Private Sub EditFolder()
    On Error GoTo Err_EditFolder

    Dim sPath As String
    Dim sPrompt As String
    Dim sTitle As String
    
    ' set default properties
    sPath = m_SelectedItem.Value
    sPrompt = "Select destination path"
    sTitle = "Browse for folders"
    ' raise this event for user cutomizations
    RaiseEvent BrowseForFolder(m_SelectedItem, sTitle, sPath, sPrompt)
    ' browse for folder
    If BrowseForFolder(Extender.Parent.hwnd, sPrompt, sPath) Then
        ' update property
        UpdateProperty sPath, True
    End If
    
    Exit Sub
Err_EditFolder:
    m_bBrowseMode = False
End Sub

Private Sub EditColor()
    On Error GoTo Err_EditColor
    
    Dim CurrColor As Long
    Dim dlgCMD As New cCommonDialog
    
    CurrColor = Val(m_SelectedItem.Value)
    With dlgCMD
        .DialogTitle = m_SelectedItem.Caption
        .CancelError = True
        .flags = CC_AnyColor Or CC_FullOpen 'CC_RGBInit
        .Color = CurrColor
        .hwnd = hwnd
        .ShowColor
        UpdateProperty .Color, True
    End With
    ' clean it up
    Set dlgCMD = Nothing
    
    Exit Sub
Err_EditColor:
    Set dlgCMD = Nothing
    m_bBrowseMode = False
End Sub

Private Sub ShowTextBox()
    On Error GoTo Err_ShowTextBox
    
    Dim strDisplayStr As String
    Dim MinValue As Variant
    Dim MaxValue As Variant
    
    txtBox.Visible = False
    ' use correct Font
    SetControlFont txtBox
    ' update dimensions
    txtBox.Left = rc.Left + (2 * Screen.TwipsPerPixelX)
    txtBox.Top = rc.Top + Screen.TwipsPerPixelY
    txtBox.Width = rc.Width - (2 * Screen.TwipsPerPixelX)
    txtBox.Height = rc.Height - (2 * Screen.TwipsPerPixelY)
    If HasGraphicInterface(m_SelectedItem) Then
        txtBox.Left = rc.InterfaceLeft
        txtBox.Width = rc.Width - (rc.InterfaceLeft - rc.Left)
    End If
    ' check for max length specification
    m_SelectedItem.GetRange MinValue, MaxValue
    ' if max value is numeric then we have a length restriction
    If IsNumeric(MaxValue) Then
        txtBox.MaxLength = MaxValue
    Else
        txtBox.MaxLength = 255
    End If
    If m_SelectedItem.Format = "Password" Then
        txtBox.PasswordChar = "*"
        strDisplayStr = m_SelectedItem.Value
    Else
        txtBox.PasswordChar = ""
        strDisplayStr = GetDisplayString(m_SelectedItem)
    End If
    txtBox.Text = strDisplayStr
    If IsReadOnly(m_SelectedItem) Then
        txtBox.Locked = True
    Else
        txtBox.Locked = False
    End If
    txtBox.Enabled = True
    txtBox.Visible = True

    Exit Sub
Err_ShowTextBox:
    ' cant show the text box within the current cell
    txtBox.Visible = False
End Sub

Private Sub ShowBrowseButton()
    On Error GoTo Err_ShowBrowseButton
    ' configure button dimensions
    With cmdBrowse
        .Top = rc.ButtonTop
        .Width = rc.ButtonWidth
        .Left = rc.ButtonLeft
        .Height = rc.ButtonHeight
    End With
    ' update button image
    If m_SelectedItem.ValueType = psFile Or _
       m_SelectedItem.ValueType = psFolder Or _
       m_SelectedItem.ValueType = psColor Or _
       m_SelectedItem.ValueType = psPicture Or _
       m_SelectedItem.ValueType = psCustom Or _
       m_SelectedItem.ValueType = psFont Then
        cmdBrowse.Picture = StdImages.ListImages("dots").Picture
    Else
        cmdBrowse.Picture = StdImages.ListImages("drop").Picture
    End If
    cmdBrowse.Visible = True
    txtBox.Width = (txtBox.Width - cmdBrowse.Width) + Screen.TwipsPerPixelX

    Exit Sub
Err_ShowBrowseButton:
    txtBox.Visible = False
End Sub

Private Sub ShowUpDown()
    On Error GoTo Err_ShowBrowseButton

    Dim MinValue As Variant
    Dim MaxValue As Variant
    Dim Increment As Variant
    
    ' get property range
    m_SelectedItem.GetRange MinValue, MaxValue
    ' set default values
    If IsEmpty(MinValue) Then MinValue = -999999
    If IsEmpty(MaxValue) Then MaxValue = 999999
    ' update min/max values
    UpDown.Min = MinValue
    UpDown.Max = MaxValue
    ' add increment
    Increment = m_SelectedItem.UpDownIncrement
    ' check for a numeric value here
    If IsNumeric(Increment) Then
        If Increment = 0 Then Increment = 1
    Else
        Increment = 1
    End If
    ' configure updown dimensions
    With UpDown
        .Increment = Increment
        .Top = rc.ButtonTop
        .Width = rc.ButtonWidth
        .Left = rc.ButtonLeft
        .Height = rc.ButtonHeight
        .Visible = True
    End With
    txtBox.Width = (txtBox.Width - UpDown.Width) + Screen.TwipsPerPixelX
    
    Exit Sub
Err_ShowBrowseButton:
    txtBox.Visible = False
End Sub

Friend Function GetRowObject(ByVal Row As Integer) As Object
    On Error GoTo Err_GetRowObject

    Dim Ptr As Long
    
    If fGrid.Rows = 0 Then Exit Function
    Ptr = fGrid.RowData(Row)
    Set GetRowObject = ObjectFromPtr(Ptr)
    Exit Function

Err_GetRowObject:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "GetRowObject")
End Function

Private Sub CollapseCategory(objCat As TCategory)
    On Error GoTo Err_CollapseCategory

    Dim Cancel As Boolean
    Dim i As Integer
    Dim Row As Integer
    Dim Cell As Integer
    
    If ((objCat.Expanded = False) Or (m_ExpandableCategories = False)) Then Exit Sub
    Cancel = False
    RaiseEvent CategoryCollapsed(Cancel)
    If Cancel = True Then Exit Sub
    Row = m_SelectedRow
    StopFlicker hwnd
    'fGrid.Redraw = False
    Cell = Cell_Save
    With fGrid
        For i = 1 To objCat.Properties.Count
            .RowHeight(objCat.Properties(i).Row) = 0
            .Row = objCat.Properties(i).Row
            If m_hIml <> 0 Then
                Cell_ClearPicture objCat.Properties(i).Row, colName
                Cell_ClearPicture objCat.Properties(i).Row, colValue
            End If
        Next
    End With
    SetState Row, ColStatus, False
    objCat.Expanded = False
' 显示分类图像BUG 季祝建 2008-04-21
    ' set the apropriate icon
'    objCat.Image = m_CollapsedImage
'    Cell_DrawPicture Row, colName, m_CollapsedImage
    Grid_Resize
    Cell_Restore Cell
    'fGrid.Redraw = True
    Release
    Exit Sub
Err_CollapseCategory:
    'fGrid.Redraw = True
    Release
End Sub

Private Sub ExpandCategory(objCat As TCategory)
    On Error GoTo Err_ExpandCategory

    Dim Cancel As Boolean
    Dim i As Integer
    Dim Row As Integer
    Dim objProp As TProperty
    
    If ((objCat.Expanded = True) Or (m_ExpandableCategories = False)) Then Exit Sub
    Cancel = False
    RaiseEvent CategoryExpanded(Cancel)
    If Cancel = True Then Exit Sub
    Row = m_SelectedRow
    'fGrid.Redraw = False
    StopFlicker hwnd
    For i = 1 To objCat.Properties.Count
        fGrid.RowHeight(objCat.Properties(i).Row) = DefaultHeight
        If m_hIml <> 0 Then
            Cell_DrawPicture objCat.Properties(i).Row, colName, objCat.Properties(i).Image
        End If
        If HasGraphicInterface(objCat.Properties(i)) Then
            DrawGraphicInterface objCat.Properties(i), objCat.Properties(i).Row
        End If
    Next
    SetState Row, ColStatus, True
    objCat.Expanded = True
' 显示分类图像BUG 季祝建 2008-04-21
    ' set the apropriate icon
'    objCat.Image = m_ExpandedImage
'    Cell_DrawPicture Row, colName, m_ExpandedImage
    Grid_Resize
    'fGrid.Redraw = True
    Release
    Exit Sub
Err_ExpandCategory:
    'fGrid.Redraw = True
    Release
    Err.Raise Err.Number, GenErrSource(m_constClassName, "ExpandCategory")
End Sub

Private Sub SetState( _
    Row As Integer, _
    Col As Integer, _
    bExpanded As Boolean)
       
    On Error GoTo Err_SetState
    
    Dim Cell As Integer
    
    With fGrid
        .Redraw = False
        Cell = Cell_Save
        .Row = Row
        .Col = Col
        .CellPictureAlignment = flexAlignCenterCenter
        If bExpanded = False Then
            Set .CellPicture = StdImages.ListImages("plus").Picture
        Else
            Set .CellPicture = StdImages.ListImages("minus").Picture
        End If
        Cell_Restore Cell
        .Redraw = True
    End With
    
    Exit Sub
Err_SetState:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "SetState")
End Sub

Private Function IsWindowLess(objProp As Object) As Boolean
    On Error GoTo Err_IsWindowLess

    Select Case objProp.ValueType
        Case 0, psFont To psCombo, psLongText To psDropDownCheckList, psObject, psDate, psBoolean, psCustom
            If objProp.ValueType = psBoolean And objProp.Format = "checkbox" Then
                IsWindowLess = True
            Else
                IsWindowLess = False
            End If
        Case Else
            IsWindowLess = True
    End Select
    Exit Function

    Exit Function
Err_IsWindowLess:
    IsWindowLess = True
End Function

Private Function GetNextVisibleRow() As Integer
    On Error GoTo Err_GetNextVisibleRow

    Dim Row As Integer
    Dim obj As Object
    
    ' get the next row
    Row = fGrid.Row + 1
    ' loop to get a property row
    Do While fGrid.RowHeight(Row) = 0 And Row < fGrid.Rows + 1
        Row = Row + 1
    Loop
    ' return property row
    If Row <= fGrid.Rows Then
        GetNextVisibleRow = Row
    Else
        GetNextVisibleRow = -1
    End If

    Exit Function
Err_GetNextVisibleRow:
    GetNextVisibleRow = -1
End Function

Private Function GetPreviousVisibleRow() As Integer
    On Error GoTo Err_GetPreviousVisibleRow

    Dim Row As Integer
    
    ' get previous row
    Row = fGrid.Row - 1
    ' loop for a property
    Do While fGrid.RowHeight(Row) = 0 And Row > -1
        Row = Row - 1
    Loop
    ' return new row
    GetPreviousVisibleRow = Row

    Exit Function
Err_GetPreviousVisibleRow:
    GetPreviousVisibleRow = -1
End Function

Friend Sub DoSort( _
       RowStart As Integer, _
       RowEnd As Integer, _
       Optional Col As Integer, _
       Optional SortMethod As Integer = 1)

    On Error GoTo Err_DoSort
    
    Dim Cell As Integer
    
    With fGrid
        .Redraw = False
        Cell = Cell_Save
        .Row = RowStart
        .Col = Col
        .RowSel = RowEnd
        .ColSel = Col         ' fGrid.Cols - 1
        .Sort = SortMethod  ' 1 - Generic ascending.
        Cell_Restore Cell
        .Redraw = True
    End With
end_sort:

    Exit Sub
Err_DoSort:
End Sub

Private Sub Hilite(ByVal Row As Integer)
    On Error GoTo Err_Hilite
    
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR
    
    ' dehilite current row
    DeHilite
    ' save selected row
    m_SelectedRow = Row
    ' get the object associated with this row
    Set m_SelectedItem = GetRowObject(Row)
    ' nothing found then exit
    If m_SelectedItem Is Nothing Then Exit Sub
    ' activate selection
    m_SelectedItem.Selected = True
    With fGrid
        .Redraw = False
        .Row = Row                      ' set grid row
        .Col = colName                        ' set grid color col #2
        .CellBackColor = m_SelBackColor
        .CellForeColor = m_SelForeColor
        
        tmpBackColor = m_BackColor
        tmpForeColor = vbBlack
        If IsProperty(m_SelectedItem) Then
            If m_SelectedItem.ReadOnly = True Then
                tmpForeColor = m_SelectedItem.ForeColor
            End If
        End If
        .Col = colValue                        ' set grid color col #2
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForeColor
        
'        .CellBackColor = vbWhite
'        .CellForeColor = vbBlack
'        If IsProperty(m_SelectedItem) Then
'            If HasGraphicInterface(m_SelectedItem) Then
'                DrawGraphicInterface m_SelectedItem, Row
'            End If
'        End If
        ' save cell dimensions
        StoreCellPosition
        Cell_DrawPicture Row, colName, m_SelectedItem.Image
        .Redraw = True
    End With
    Exit Sub
Err_Hilite:
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Hilite")
End Sub

Private Sub DeHilite()
    On Error GoTo Err_DeHilite

    Dim obj As Object
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR
    
    ' finds the row associate with the selected object
    Set obj = GetRowObject(m_SelectedRow)
    ' row not found then exit
    If obj Is Nothing Then Exit Sub
    With fGrid
        .Redraw = False
        .Row = m_SelectedRow            ' set grid row
        .Col = colName                        ' set grid col
        .ColSel = colValue
        obj.Selected = False
        GetObjectColors obj, tmpBackColor, tmpForeColor
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForeColor
'        If HasGraphicInterface(m_SelectedItem) Then
'            DrawGraphicInterface m_SelectedItem, m_SelectedRow
'        End If
        Cell_DrawPicture m_SelectedRow, colName, obj.Image
        .Redraw = True
    End With
    Exit Sub

Err_DeHilite:
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "DeHilite")
End Sub

Friend Function FindGridRow(obj As Object) As Integer
    On Error GoTo Err_FindGridRow

    Dim i As Integer
    Dim Ptr As Long
    
    Ptr = obj.Handle
    For i = 0 To fGrid.Rows - 1
        If fGrid.RowData(i) = Ptr Then
            FindGridRow = i
            Exit Function
        End If
    Next
    FindGridRow = -1
    Exit Function

Err_FindGridRow:
    FindGridRow = -1
End Function

Private Sub Cell_DrawPicture( _
        Row As Integer, _
        Col As Integer, _
        Image As Variant)
    
    On Error GoTo Err_Cell_DrawPicture
    
    Dim BkColor As OLE_COLOR
    Dim obj As Object
    
    Set obj = GetRowObject(Row)
    If obj Is Nothing Then Exit Sub
    GetObjectColors obj, BkColor
    Cell_DrawPictureEx Row, Col, Image, m_hIml, BkColor
    Exit Sub

Err_Cell_DrawPicture:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Cell_DrawPicture")
End Sub

Private Sub Cell_ClearPicture( _
    Row As Integer, _
    Col As Integer)
    
    On Error GoTo Err_Cell_ClearPicture
    
    Dim obj As Object
    Dim tmpBackColor As OLE_COLOR
    Dim Cell As Integer
        
    Set obj = GetRowObject(Row)
    If obj Is Nothing Then Exit Sub
    With fGrid
        .Redraw = False
        Cell = Cell_Save
        .Row = Row
        .Col = Col
        Set .CellPicture = Nothing
        Cell_Restore Cell
        .Redraw = True
    End With

    Exit Sub
Err_Cell_ClearPicture:
End Sub

Private Function IsProperty(ByVal obj As Object) As Boolean
    IsProperty = TypeName(obj) = "TProperty"
End Function

Private Function GetDisplayString(objProp As TProperty) As String
    Dim strFormat As String
    Dim bUseDefault As Boolean
    Dim strDisplayStr As String
    
    strFormat = objProp.Format
    If (objProp.ValueType = psCustom) Or (strFormat = "CustomDisplay") Then
        bUseDefault = False
        RaiseEvent GetDisplayString(objProp, strDisplayStr, bUseDefault)
    Else
        bUseDefault = True
    End If
    If bUseDefault = True Then
        strDisplayStr = GetDefaultDisplayString(objProp)
    End If
    If objProp.Format = "Password" Then
        strDisplayStr = String(Len(strDisplayStr), "*")
    End If
    GetDisplayString = strDisplayStr
End Function

Private Function GetDefaultDisplayString(objProp As TProperty) As String
    On Error GoTo Err_GetDefaultDisplayString

    Dim strDisplayStr As String
    Dim i As Integer
    Dim strTemp As String
    Dim lsValue As TListValue
    
    Select Case objProp.ValueType
        Case psDropDownCheckList
            strDisplayStr = "(List)"
        Case psFont
            If IsObject(objProp.Value) Then
                If Not objProp.Value Is Nothing Then
                    If objProp.Format <> "" Then
                        strDisplayStr = FormatFont(objProp.Value, objProp.Format)
                    Else
                        strDisplayStr = objProp.Value.Name
                    End If
                Else
                    strDisplayStr = "(None)"
                End If
            End If
        Case psObject
            strDisplayStr = "(Object)"
        Case psLongText
            strDisplayStr = "(Text)"
        Case psPicture
            strDisplayStr = "(Picture)"
        Case psTime
            strDisplayStr = Format(objProp.Value, "hh:mm:ss")
        Case psColor
            If objProp.Format <> "" Then
                strDisplayStr = FormatColor(objProp.Value, objProp.Format)
            Else
                strDisplayStr = objProp.Value
            End If
        
        Case psCombo
            For Each lsValue In objProp.ListValues
                If objProp.Value = lsValue.Value Then
                    strDisplayStr = lsValue.Caption
                    GoTo Exit_GetDefaultDisplayString
                End If
            Next
            ' not found then set the default value
            strDisplayStr = objProp.Value
        
        Case psDropDownList
            
            For i = 1 To objProp.ListValues.Count
                Set lsValue = objProp.ListValues(i)
                If lsValue.Value = objProp.Value Then
                    strDisplayStr = lsValue.Caption
                    GoTo Exit_GetDefaultDisplayString
                End If
            Next
            
        Case Else
            If IsArray(objProp.Value) Then
                strDisplayStr = "(Array)"
            Else
                If objProp.Format <> "checkbox" And objProp.Format <> "Password" And objProp.Format <> "" And objProp.Format <> "CustomDisplay" Then
                    strDisplayStr = Format(objProp.Value, objProp.Format)
                Else
                    If objProp.ListValues.Count > 0 Then
                        For i = 1 To objProp.ListValues.Count
                            If objProp.Value = objProp.ListValues(i).Value Then
                                strDisplayStr = objProp.ListValues(i).Caption
                                GoTo Exit_GetDefaultDisplayString
                            End If
                        Next
                    End If
                    strDisplayStr = objProp.Value
                End If
            End If
    End Select

Exit_GetDefaultDisplayString:
    GetDefaultDisplayString = strDisplayStr

    Exit Function
Err_GetDefaultDisplayString:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "GetDefaultDisplayString")
End Function

Private Sub ToggleCategoryState()
    If m_SelectedItem.Expanded = True Then
        CollapseCategory m_SelectedItem
    Else
        ExpandCategory m_SelectedItem
    End If
End Sub

Private Function Cell_Save() As Integer
    Cell_Save = MakeWord(fGrid.Row, fGrid.Col)
End Function

Private Function Cell_Restore(Info As Integer)
    On Error Resume Next
    fGrid.Row = HiByte(Info)
    fGrid.Col = LoByte(Info)
End Function

Private Function IsIncremental(Prop As TProperty) As Boolean
    Dim MinValue As Variant
    Dim MaxValue As Variant
    Prop.GetRange MinValue, MaxValue
    IsIncremental = _
       (Not IsEmpty(MinValue) Or _
       Not IsEmpty(MaxValue) Or _
       Prop.UpDownIncrement > 0) And _
       (IsWindowLess(Prop) And _
       Prop.ValueType <> psString)
End Function

Private Sub StoreCellPosition()
    On Error Resume Next
    Dim edgeX As Single
    Dim edgeY As Single
    If BorderStyle = psBorderSingle Then
        edgeX = 2 * Screen.TwipsPerPixelX
        edgeY = 2 * Screen.TwipsPerPixelY
    End If
    With fGrid
        ' set column to the value column
        .Col = colValue
        ' get cell rect
        rc.Left = .CellLeft + Screen.TwipsPerPixelX
        rc.Top = .CellTop
        rc.Height = .CellHeight
        rc.Width = .CellWidth
        ' button properties
        rc.ButtonTop = rc.Top
        rc.ButtonWidth = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
        rc.ButtonLeft = (rc.Left + rc.Width) - (rc.ButtonWidth)
        rc.ButtonHeight = rc.Height
        ' get browse window initial rect
        rc.WindowLeft = (Extender.Left + .CellLeft) + edgeX
'        RC.WindowLeft = .CellLeft '+ edgeX
        rc.WindowTop = Extender.Top + .CellTop + .CellHeight + edgeY
'        RC.WindowTop = .CellTop + .CellHeight '+ edgeY
        rc.WindowWidth = rc.Width
        ' interface left is 16 pixels
        'rc.InterfaceLeft = .CellLeft + edgeX + TextWidth(Space(m_lPadding)) '(16 * Screen.TwipsPerPixelX)
        rc.InterfaceLeft = .CellLeft + edgeX + (m_lPadding * TextWidth(" "))
    End With
End Sub

Private Sub Grid_ShowCategories()
On Error GoTo Err_Grid_ShowCategories

    Dim Row As Integer
    Dim Cell As Integer
    Dim objCat As TCategory
    Dim i As Integer
    Dim j As Integer
    Dim Prop As TProperty
    Dim objTemp As Object
    Dim obj As Object
    Dim Handle As Long
    
    StopFlicker hwnd
    ' get the object related with this row
    Set objTemp = GetRowObject(m_SelectedRow)
    ' hide all the controls
    HideControls
    ' save cell position
    Cell = Cell_Save
    ' if it is to disable categories...
    If m_ShowCategories = False Then
        ' set column #0 width to 0
        fGrid.ColWidth(ColStatus) = 0
        If m_Categories.Count > 0 Then
            For i = 1 To m_Categories.Count
                Row = m_Categories(i).Row
                fGrid.RowHeight(Row) = 0
                ' clear minus/plus picture
                Cell_ClearPicture Row, ColStatus
                For j = 1 To m_Categories(i).Properties.Count
                    Row = m_Categories(i).Properties(j).Row
                    fGrid.RowHeight(Row) = DefaultHeight
                Next
            Next
            ' sort entire row count
            DoSort 0, fGrid.Rows - 1, colName, flexSortStringNoCaseAsending
        End If
    Else
        ' set column #0 width to 0
        fGrid.ColWidth(ColStatus) = COL_WIDTH * Screen.TwipsPerPixelX
        For i = 1 To m_Categories.Count
            Row = m_Categories(i).Row
            fGrid.RowHeight(Row) = DefaultHeight
            SetState Row, ColStatus, m_Categories(i).Expanded
        Next
        ' sort entire row count
        DoSort 0, fGrid.Rows - 1, colSort, flexSortNumericAscending
    End If
    Grid_Reindex
    ' select the object row
    If Not objTemp Is Nothing Then
        m_SelectedRow = objTemp.Row
    Else
        If fGrid.Rows > 0 Then
            m_SelectedRow = 0
        End If
    End If
    ' restore position
    Cell_Restore Cell
    ' resize the grid
    Grid_Resize
    Release
    
    Exit Sub
Err_Grid_ShowCategories:
    Release
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_ShowCategories")
End Sub

Private Sub Grid_Reindex()
    Dim Row As Integer
    Dim obj As Object
    
    For Row = 0 To fGrid.Rows - 1
        Set obj = GetRowObject(Row)
        If Not obj Is Nothing Then
            obj.Row = Row
        End If
    Next
End Sub

Private Sub Grid_Index()
    Dim Handle As Long
    Dim Row As Integer
    Dim i As Integer
    Dim j As Integer
    
    If m_Categories.Count > 0 Then
        For i = 1 To m_Categories.Count
            Row = m_Categories(i).Row
            Handle = MakeDWord(m_Categories(i).Index, 0)
            fGrid.TextMatrix(Row, colSort) = Handle
            For j = 1 To m_Categories(i).Properties.Count
                Row = m_Categories(i).Properties(j).Row
                Handle = MakeDWord(m_Categories(i).Index, m_Categories(i).Properties(j).Index)
                fGrid.TextMatrix(Row, colSort) = Handle
            Next
        Next
    End If
End Sub

Private Sub SetControlFont(Ctl As Control)
    Ctl.FontName = m_Font.Name
    Ctl.FontSize = m_Font.Size
End Sub

Private Function FixTopPos(lHeight) As Long
    On Error Resume Next
    lHeight = CLng(lHeight)
    If rc.WindowTop + lHeight + 300 < UserControl.Extender.Parent.ScaleHeight Then
        FixTopPos = rc.WindowTop
    Else
        FixTopPos = rc.WindowTop - (rc.Height + lHeight)
    End If
End Function

Private Sub HideBrowseWnd()
    If Not m_BrowseWnd Is Nothing Then
        If m_BrowseWnd.Visible = True Then
            m_BrowseWnd.Visible = False
        End If
    End If
    m_bBrowseMode = False
End Sub

Private Sub DrawCheckBox( _
    Row As Integer, _
    objProp As TProperty)
    
    Dim Image As String
    Dim BkColor As OLE_COLOR
    
    If objProp.Value = True Then
        Image = "check_on"
    Else
        Image = "check_off"
    End If
    If objProp.Selected Then
        BkColor = vbWhite
    Else
        If objProp.BackColor = CLR_INVALID Then
            BkColor = m_BackColor
        Else
            BkColor = objProp.BackColor
        End If
    End If
    Cell_DrawPictureEx Row, colValue, Image, m_hImlStd, BkColor
End Sub

Private Sub DrawColorBox(Row As Integer, _
       objProp As TProperty)
    Dim Image As String
    Dim BkColor As OLE_COLOR
    
    Image = "frame"
    BkColor = objProp.Value
    Cell_DrawPictureEx Row, colValue, StdImages.ListImages(Image).Index, m_hImlStd, BkColor
End Sub

Private Sub Cell_DrawPictureEx( _
       ByVal Row As Integer, _
       ByVal Col As Integer, _
       ByVal Image As Variant, _
       hIml As Long, _
       Optional BkColor As OLE_COLOR = 0)
    
    On Error GoTo Err_Cell_DrawPictureEx
    
    Dim obj As Object
    
    If fGrid.RowHeight(Row) = 0 Or hIml = 0 Then Exit Sub
    With fGrid
        .Redraw = False
        .Row = Row
        .Col = Col
        .CellPictureAlignment = flexAlignLeftCenter
        If BkColor <> -1 Then
            Image_List(hIml).BackColor = BkColor
        End If
        Set .CellPicture = Image_List(hIml).Overlay(Image, Image)
        .Redraw = True
    End With
    
    Exit Sub
Err_Cell_DrawPictureEx:
End Sub

Private Sub RecalcPadding()
    Dim Twips As Long
    Dim w As Single
    
    If Not m_Font Is Nothing Then
        Twips = 17 * Screen.TwipsPerPixelX
        Set UserControl.Font = m_Font
        w = Twips / TextWidth(" ")
        m_lPadding = w
    End If
    Grid_Resize
End Sub

Private Function HasGraphicInterface(objProp As Object) As Boolean
    If objProp Is Nothing Or IsProperty(objProp) = False Then
        HasGraphicInterface = False
        Exit Function
    End If
    HasGraphicInterface = (objProp.ValueType = psBoolean And objProp.Format = "checkbox") Or _
       objProp.ValueType = psColor
End Function

Private Sub DrawGraphicInterface(objProp As TProperty, Row As Integer)
    'Dim Row As Integer
    
    'Row = FindGridRow(objProp)
    'If Row = -1 Then Exit Sub
    If (objProp.ValueType = psBoolean And objProp.Format = "checkbox") Then
        DrawCheckBox Row, objProp
    Else
        DrawColorBox Row, objProp
    End If
End Sub

Private Sub SelectText()
    On Error Resume Next
    txtBox.SetFocus
    txtBox.SelStart = 0
    txtBox.SelLength = Len(txtBox.Text)
    '    Call SendMessage(txtBox.hWnd, EM_SETSEL, 0, Len(txtBox.Text))
End Sub

Public Sub LoadFromFile(ByVal Filename As String, _
       ByVal Section As String)
    On Error GoTo Err_LoadFromFile

    'StopFlicker hwnd
    Dim Col As Collection
    Set Col = EnumSections(Filename, Section)
    If Col Is Nothing Then
        'Release
        Exit Sub
    End If
    On Error Resume Next
    Set Font = ReadProperty(Col("Font"), Ambient.Font)
    Set CatFont = ReadProperty(Col("CatFont"), Ambient.Font)
'    AllowEmptyValues = ReadProperty(Col("AllowEmptyValues"), m_def_AllowEmptyValues)
    ExpandableCategories = ReadProperty(Col("ExpandableCategories"), m_def_ExpandableCategories)
    NameWidth = ReadProperty(Col("NameWidth"), m_def_NameWidth)
    RequiresEnter = ReadProperty(Col("RequiresEnter"), m_def_RequiresEnter)
    ShowToolTips = ReadProperty(Col("ShowToolTips"), m_def_ShowToolTips)
    CatBackColor = ReadProperty(Col("CatBackColor"), m_def_CatBackColor)
    CatForeColor = ReadProperty(Col("CatForeColor"), m_def_CatForeColor)
    SelBackColor = ReadProperty(Col("SelBackColor"), m_def_SelBackColor)
    SelForeColor = ReadProperty(Col("SelForeColor"), m_def_SelForeColor)
    BackColor = ReadProperty(Col("BackColor"), m_def_BackColor)
    ForeColor = ReadProperty(Col("ForeColor"), m_def_ForeColor)
    GridColor = ReadProperty(Col("GridColor"), m_def_GridColor)
    BorderStyle = ReadProperty(Col("BorderStyle"), 1)
    Appearance = ReadProperty(Col("Appearance"), 1)
    'Release
    Exit Sub
Err_LoadFromFile:
    'Release
    Err.Raise Err.Number, GenErrSource(m_constClassName, "LoadFromFile")
End Sub

Private Function ReadProperty(Prop As Variant, DefProp As Variant) As Variant
    Dim objTemp As Object
    If IsVarEmpty(Prop) Then
        If IsObject(DefProp) Then
            Set ReadProperty = DefProp
        Else
            ReadProperty = DefProp
        End If
    Else
        If IsObject(DefProp) Then
            If TypeOf DefProp Is StdFont Then
                Set objTemp = FontFromStr(Prop)
                Set ReadProperty = objTemp
            Else
                Set ReadProperty = Prop
            End If
        Else
            ReadProperty = Prop
        End If
    End If
End Function

Public Sub SaveToFile( _
       ByVal Filename As String, _
       ByVal Section As String)

    On Error GoTo Err_SaveFile

    Dim i As Integer
    Dim j As Integer
    Dim hFile As Integer
    
    'StopFlicker hwnd
    m_strText = "[" & Section & "]" & vbCrLf
    Call WriteProperty("Font", StrFromFont(m_Font))
    Call WriteProperty("CatFont", StrFromFont(m_CatFont))
'    Call WriteProperty("AllowEmptyValues", m_AllowEmptyValues)
    Call WriteProperty("ExpandableCategories", m_ExpandableCategories)
    Call WriteProperty("NameWidth", m_NameWidth)
    Call WriteProperty("RequiresEnter", m_RequiresEnter)
    Call WriteProperty("ShowCategories", m_ShowCategories)
    Call WriteProperty("ShowToolTips", m_ShowToolTips)
    Call WriteProperty("CatBackColor", m_CatBackColor)
    Call WriteProperty("CatForeColor", m_CatForeColor)
    Call WriteProperty("SelBackColor", m_SelBackColor)
    Call WriteProperty("SelForeColor", m_SelForeColor)
    Call WriteProperty("BackColor", m_BackColor)
    Call WriteProperty("ForeColor", m_ForeColor)
    Call WriteProperty("GridColor", m_GridColor)
    Call WriteProperty("BorderStyle", BorderStyle)
    Call WriteProperty("Appearance", Appearance)
    Call WriteProperty("ExpandedImage", m_ExpandedImage)
    Call WriteProperty("CollapsedImage", m_CollapsedImage)
    hFile = FreeFile
    Open Filename For Output As #hFile
    Print #hFile, m_strText
    Close #hFile
    'Release

    Exit Sub
Err_SaveFile:
    Close #hFile
    'Release
    Err.Raise Err.Number, GenErrSource(m_constClassName, "SaveFile")
End Sub

Private Sub WriteProperty(ByVal Prop As String, ByVal Value As Variant)
    m_strText = m_strText & Prop & "=" & Value & vbCrLf
End Sub

Private Function DefaultHeight() As Long
    ' return height in pixels
    DefaultHeight = 18 * Screen.TwipsPerPixelY
End Function

Private Function IsReadOnly(Prop As TProperty) As Boolean
    With Prop
        IsReadOnly = _
           .ValueType = psLongText Or _
           .ValueType = psPicture Or _
           .ValueType = psDropDownCheckList Or _
           .ValueType = psDropDownList Or _
           .ValueType = psBoolean Or _
           .ValueType = psFont Or _
           .ReadOnly Or _
           IsArray(.Value)
    End With
End Function

Private Function IsBrowsable(Prop As TProperty) As Boolean
    With m_SelectedItem
        IsBrowsable = _
           .ValueType = psColor Or _
           .ValueType = psFile Or _
           .ValueType = psFolder Or _
           .ValueType = psPicture Or _
           .ValueType = psFont Or _
           .ValueType = psDate
    End With
End Function

Private Function Pad(ByVal Text As String) As String
    If m_hIml = 0 Then
        Pad = Text
    Else
        Pad = Space(m_lPadding) & Text
    End If
End Function

'Private Function AvgTextWidth() As Single
'    Dim avgWidth As Single
'    ' Get the average character width of the current list box Font
'    ' (in pixels) using the form's TextWidth width method.
'    avgWidth = UserControl.TextWidth("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
'    avgWidth = avgWidth / 52
'    ' Set the white space you want between columns.
'    AvgTextWidth = avgWidth
'End Function

' 增加对双字节的支持 季祝建 2008-04-21
Private Function EvaluateTextWidth(ByVal s As String) As Single
'    Dim avgWidth As Single
    
'    avgWidth = AvgTextWidth
'    EvaluateTextWidth = Len(Trim(s)) * avgWidth
    EvaluateTextWidth = IIf(m_hIml <> 0, 200, 0) + lstrlen(Trim(s)) * 100 + 50 ' 图像+字符+边框
End Function

Private Function Cell_NameWidth() As Single
    Dim tW As Single
    Dim nw As Single
    Dim Index As Integer
    
    nw = 0
    ' calculate the automatic name width
    For Index = 1 To m_Properties.Count
        tW = EvaluateTextWidth(Properties(Index).Caption)
        If tW > nw Then
            nw = tW
        End If
    Next Index
    Cell_NameWidth = nw
End Function

Private Sub Grid_Paint()
On Error GoTo Err_Grid_Paint
    
    Dim Cat As Integer
    Dim Prop As Integer
    Dim RowStart As Integer
    Dim RowSel As Integer
    Dim Cell As Integer
    Dim Row As Integer
    Dim obj As Object
    
    RecalcPadding
    ' check for Categories. If no categories
    ' is found then clear grid and exit
    If m_Categories.Count = 0 Then
        Grid_Clear
        Exit Sub
    End If
    ' hide all visible control
    HideControls
    ' update category order
    Grid_ShowCategories
    ' draw grid cells
    With fGrid
        ' to avoid flickering
        .Redraw = False
        ' save current cell
        Cell = Cell_Save
        .GridColor = m_GridColor
        .GridColorFixed = m_BackColor
        .BackColorFixed = m_BackColor
        .BackColorSel = m_SelBackColor
        .BackColorBkg = m_BackColor
        .BackColorUnpopulated = m_BackColor
        .BackColor = m_BackColor
        .ForeColorSel = m_SelForeColor
        .ForeColor = m_ForeColor
        Set .Font = m_Font
        ' restore cell position
        Cell_Restore Cell
        .Redraw = True
    End With

    Exit Sub
Err_Grid_Paint:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "Grid_Paint")
End Sub

Friend Sub TriggerEvent(ByVal RaisedEvent As String, ParamArray aParams())
    Select Case RaisedEvent
        Case "CaptionChanged"
            Cell_CaptionChanged aParams(0), aParams(1), aParams(2)
        Case "ValueChanged"
            Cell_ValueChanged aParams(0), aParams(1)
        Case "AddNewCategory"
            AddNewCategory aParams(0)
        Case "AddNewProperty"
            AddNewProperty aParams(0), aParams(1)
        Case "SelectedChanged"
            'DeHilite
            'Hilite aParams(1)
        Case "ForeColorChanged"
            Cell_ChangeForeColor aParams(1), aParams(2)
        Case "BackColorChanged"
            Cell_ChangeBackColor aParams(1), aParams(2)
    End Select
End Sub

Private Sub Cell_ChangeForeColor(ByVal Row As Integer, ByVal New_Color As OLE_COLOR)
    Dim Cell As Integer
    Cell = Cell_Save
    With fGrid
        .Row = Row
        .Col = colName
        .ColSel = colValue
        If m_SelectedRow <> Row Then
            fGrid.CellForeColor = New_Color
        Else
            fGrid.CellForeColor = m_SelForeColor
        End If
    End With
    Cell_Restore Cell
End Sub

Private Sub Cell_ChangeBackColor(ByVal Row As Integer, ByVal New_Color As OLE_COLOR)
    Dim Cell As Integer
    Cell = Cell_Save
    With fGrid
        .Row = Row
        .Col = colName
        .ColSel = colValue
        If m_SelectedRow <> Row Then
            fGrid.CellBackColor = New_Color
        Else
            fGrid.CellBackColor = m_SelBackColor
        End If
    End With
    Cell_Restore Cell
End Sub

Private Sub Cell_ValueChanged(ByVal PropObj As TProperty, ByVal Row As Integer)
    On Error GoTo Err_ValueChanged

    Dim Cell As Integer
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR
    Dim strValue As String
    
    ' save cell pos
    Cell = Cell_Save
'    txtBox.Visible = False
    ' check for the value type
    If (PropObj.ValueType <> psDropDownCheckList) Then
        HideBrowseWnd
        txtBox.Visible = False
        If UpDown.Visible = False And cmdBrowse.Visible = False Then
            HideControls
        End If
    End If
    ' configure back color
    GetObjectColors PropObj, tmpBackColor, tmpForeColor
    If PropObj.ReadOnly Then
        tmpForeColor = PropObj.ForeColor
    End If
    strValue = GetDisplayString(PropObj)
    With fGrid
        .Row = Row
        .Col = colValue
        If PropObj.BackColor <> CLR_INVALID Then
            .CellBackColor = tmpBackColor
        End If
        If PropObj.ForeColor <> CLR_INVALID Then
            .CellForeColor = tmpForeColor
        End If
        .CellAlignment = flexAlignLeftCenter
        Set .CellPicture = Nothing
        If HasGraphicInterface(PropObj) Then
            DrawGraphicInterface PropObj, Row
            strValue = Space(m_lPadding) & strValue
        End If
        .Text = strValue
    End With
    ' restore cell properties
    Cell_Restore Cell
'    txtBox.BackColor = vbYellow
    Exit Sub
Err_ValueChanged:
    Err.Raise Err.Number, GenErrSource(m_constClassName, "ValueChanged")
End Sub

Private Sub Cell_CaptionChanged(ByVal PropObj As TProperty, ByVal Row As Integer, ByVal NewCaption As String)
    On Error GoTo Err_ValueChanged

    Dim Cell As Integer
    Dim tmpBackColor As OLE_COLOR
    Dim tmpForeColor As OLE_COLOR

    ' disable drawing
    fGrid.Redraw = False
    ' save cell pos
    Cell = Cell_Save
    ' check for the value type
    If (PropObj.ValueType <> psDropDownCheckList) Then
        HideBrowseWnd
        txtBox.Visible = False
        If UpDown.Visible = False And cmdBrowse.Visible = False Then
            HideControls
        End If
    End If
    ' configure back color
    GetObjectColors PropObj, tmpBackColor, tmpForeColor
    With fGrid
        .Row = Row
        .Col = colName
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForeColor
        .CellAlignment = flexAlignLeftCenter
        If m_hIml <> 0 Then
            If PropObj.Image <> -1 Then
                Cell_DrawPicture Row, colName, PropObj.Image
            End If
        End If
        .Text = Pad(PropObj.Caption)
    End With
    ' restore cell properties
    Cell_Restore Cell
    ' enable drawing
    fGrid.Redraw = True

    Exit Sub
Err_ValueChanged:
    fGrid.Redraw = True
    Err.Raise Err.Number, GenErrSource(m_constClassName, "ValueChanged")
End Sub
'-- end code

Public Function ConvertValue(Value As Variant, ValueType As psPropertyType) As Variant
On Error GoTo Err_ConvertValue
    Select Case ValueType
        Case psInteger
            ConvertValue = CInt(Value)
        Case psLong
            ConvertValue = CLng(Value)
        Case psSingle
            ConvertValue = CSng(Value)
        Case psDouble
            ConvertValue = CDbl(Value)
        Case psCurrency
            ConvertValue = CCur(Value)
        Case psDate
            ConvertValue = CDate(Value)
        Case psString
            ConvertValue = CStr(Value)
        Case psBoolean
            ConvertValue = CBool(Value)
        Case psDecimal
            ConvertValue = Format(Value, "##,##0.00")
        Case psByte
            ConvertValue = CByte(Value)
        Case Else
            ConvertValue = Value
    End Select

Exit Function
Err_ConvertValue:
    ConvertValue = Null
End Function

