VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TPropertySheet 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2568
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4572
   ScaleHeight     =   2568
   ScaleWidth      =   4572
   ToolboxBitmap   =   "PropertySheetA.ctx":0000
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
            Picture         =   "PropertySheetA.ctx":0312
            Key             =   "frame"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0424
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0495
            Key             =   "dots"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":058F
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0689
            Key             =   "check_off"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":079B
            Key             =   "check_on"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":08AD
            Key             =   "minus"
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstCheck 
      Appearance      =   0  'Flat
      Height          =   24
      ItemData        =   "PropertySheetA.ctx":091B
      Left            =   480
      List            =   "PropertySheetA.ctx":0922
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
      Height          =   2208
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   2472
      _ExtentX        =   4360
      _ExtentY        =   3895
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483635
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   63963137
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
      Picture         =   "PropertySheetA.ctx":0929
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.ListBox lstBox 
      Appearance      =   0  'Flat
      Height          =   216
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
            Picture         =   "PropertySheetA.ctx":0A13
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0A81
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0AF2
            Key             =   "dots"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0BEC
            Key             =   "drop"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0CE6
            Key             =   "check_on_sel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0DF8
            Key             =   "check_off_sel"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":0F0A
            Key             =   "check_on_2"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":101C
            Key             =   "check_off_2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":112E
            Key             =   "check_off"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":1240
            Key             =   "check_on"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PropertySheetA.ctx":1352
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
' *******************************************************
' Control      : TPropertySheet.Ctl
' Written By   : Marclei V Silva (MVS)
' Programmer   : Marclei V Silva (MVS) [Spnorte Consultoria de Informática]
' Date Writen  : 06/16/2000 -- 09:08:30
' Description  : PropertySheet control which show property
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
Const m_constClassName = "TPropertySheet"

Private m_lngErrNum As Long
Private m_strErrStr As String
Private m_strErrSource As String

Const ColCount = 6

' Default Property Values:
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
Const m_def_AllowEmptyValues = 0
Const m_def_ExpandableCategories = 1
Const m_def_NameWidth = 1100
Const m_def_RequiresEnter = 0
Const m_def_ShowCategories = 1
'Const m_def_ImageList = ""
Const m_def_ExpandedImage = 0
Const m_def_CollapsedImage = 0
Const m_def_Initializing = 0

' Property Variables:
'Dim m_ImageList As String
Dim m_ExpandedImage As Integer
Dim m_CollapsedImage As Integer
Dim m_CatFont As Font
Dim m_ForeColor As OLE_COLOR
Dim m_GridColor As OLE_COLOR
Dim m_BackColor As OLE_COLOR
Dim m_SelBackColor As OLE_COLOR
Dim m_SelForeColor As OLE_COLOR
Dim m_CatBackColor As OLE_COLOR
Dim m_CatForeColor As OLE_COLOR
Dim m_ShowToolTips As Boolean
Dim m_SelectedItem As Object
Dim m_Enabled As Boolean
Dim m_font As Font
Dim m_AllowEmptyValues As Boolean
Dim m_ExpandableCategories As Boolean
Dim m_NameWidth As Single
Dim m_RequiresEnter As Boolean
Dim m_ShowCategories As Boolean

' private declaration
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
'Private m_AutoNameWidth As Integer
Private m_SelectedRow As Integer
Private m_lPadding As Long
Private m_Properties As Collection
Private m_strText As String
Private m_bUserMode As Boolean
'Private m_bImageListChecked As Boolean
Private m_hIml As Long
Private m_hStdIml As Long
Private m_lIconSize As Long

' Event Declarations:
Event Browse(ByVal Left, ByVal Top, ByVal Width, ByVal Prop As TProperty)
Event CategoryCollapsed(Cancel As Boolean)
Event CategoryExpanded(Cancel As Boolean)
Event EnterEditMode(ByVal Prop As TProperty, Cancel As Boolean)
Attribute EnterEditMode.VB_Description = "Occurs when the edit control is to be shown allowing the user to edit the property"
Event GetDisplayString(ByVal Prop As TProperty, DisplayString As String, UseDefault As Boolean)
Attribute GetDisplayString.VB_Description = "Occurs when the control needs the display string of a property. This event is called only if the property has the FormatProperty set to ""CustomeDisplay"""
Event ParseString(ByVal Prop As TProperty, ByVal Text As String, UseDefault As Boolean)
Attribute ParseString.VB_Description = "Occurs when the user changes a property that has the format property set to ""CustomDisplay"""
Event PropertyChanged(ByVal Prop As TProperty, NewValue, Cancel As Boolean)
Attribute PropertyChanged.VB_Description = "Occurs when a property value is changed"
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
Event BrowseForFile(ByVal Prop As TProperty, Title As String, Filter As String, FilterIndex As Integer, flags As Long)
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
    Set Font = m_font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_font = New_Font
    PropertyChanged "Font"
    'RecalcPadding
    'Grid_ChangeFont False
    'Grid_Resize
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AllowEmptyValues() As Boolean
    AllowEmptyValues = m_AllowEmptyValues
End Property

Public Property Let AllowEmptyValues(ByVal New_AllowEmptyValues As Boolean)
    m_AllowEmptyValues = New_AllowEmptyValues
    PropertyChanged "AllowEmptyValues"
End Property

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
    'Grid_ShowCategories New_ShowCategories
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
    m_CatBackColor = New_CatBackColor
    PropertyChanged "CatBackColor"
    'Grid_ChangeBackColor New_CatBackColor, True
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H00FFFFFF&
Public Property Get CatForeColor() As OLE_COLOR
    CatForeColor = m_CatForeColor
End Property

Public Property Let CatForeColor(ByVal New_CatForeColor As OLE_COLOR)
    m_CatForeColor = New_CatForeColor
    PropertyChanged "CatForeColor"
    'Grid_ChangeForeColor New_CatForeColor, True
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000D&
Public Property Get SelBackColor() As OLE_COLOR
    SelBackColor = m_SelBackColor
End Property

Public Property Let SelBackColor(ByVal New_SelBackColor As OLE_COLOR)
    m_SelBackColor = New_SelBackColor
    PropertyChanged "SelBackColor"
    'Hilite m_SelectedRow 'Grid_ChangeSelBackColor New_SelBackColor
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H8000000E&
Public Property Get SelForeColor() As OLE_COLOR
    SelForeColor = m_SelForeColor
End Property

Public Property Let SelForeColor(ByVal New_SelForeColor As OLE_COLOR)
    m_SelForeColor = New_SelForeColor
    PropertyChanged "SelForeColor"
    '    Grid_ChangeSelForeColor New_SelForeColor
    'Hilite m_SelectedRow
    m_bDirty = True: Grid_Paint
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
    m_BackColor = New_BackColor
'    With fGrid
'        .BackColor = New_BackColor
'        .BackColorBkg = New_BackColor
'        .BackColorUnpopulated = New_BackColor
'        .GridColorFixed = New_BackColor
'        .BackColorFixed = New_BackColor
'    End With
'    Grid_ChangeBackColor New_BackColor, False
    PropertyChanged "BackColor"
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,&H
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    'Grid_ChangeForeColor New_ForeColor, False
    m_bDirty = True: Grid_Paint
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
    Set m_CatFont = New_CatFont
    PropertyChanged "CatFont"
    'Grid_ChangeFont True
    m_bDirty = True: Grid_Paint
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
    m_bDirty = True: Grid_Paint
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
    m_bDirty = True: Grid_Paint
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
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        Dim cx As Long, cy As Long
        If (ImageList_GetIconSize(m_hIml, cx, cy) <> 0) Then
            m_lIconSize = cy
        End If
    End If
'    PropertyChanged "ImageList"
    m_bDirty = True: Grid_Paint
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    m_bDirty = True
    Grid_Paint
End Sub

Private Sub txtBox_GotFocus()
    If IsProperty(m_SelectedItem) Then
        If m_SelectedItem.ValueType <> psTime Then
            SelectText
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()
    Set m_font = Ambient.Font
    ' recalc padding
    RecalcPadding
    Set m_CatFont = Ambient.Font
    m_Enabled = m_def_Enabled
    m_AllowEmptyValues = m_def_AllowEmptyValues
    m_ExpandableCategories = m_def_ExpandableCategories
    m_NameWidth = m_def_NameWidth
    m_RequiresEnter = m_def_RequiresEnter
    m_ShowCategories = m_def_ShowCategories
    m_ShowToolTips = m_def_ShowToolTips
    m_CatBackColor = m_def_CatBackColor
    m_CatForeColor = m_def_CatForeColor
    m_SelBackColor = m_def_SelBackColor
    m_SelForeColor = m_def_SelForeColor
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
    If (UserControl.Ambient.UserMode = False) Then
        m_bUserMode = False
        Set m_Properties = Nothing
        Set m_Properties = New Collection
        With m_Categories.Add("TPropertySheet")
            .Properties.Add "(Name)", UserControl.Ambient.DisplayName
            With .Properties.Add("Selected", "Value")
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
    Set m_font = PropBag.ReadProperty("Font", Ambient.Font)
    ' recalc padding
    RecalcPadding
    Set m_CatFont = PropBag.ReadProperty("CatFont", Ambient.Font)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_AllowEmptyValues = PropBag.ReadProperty("AllowEmptyValues", m_def_AllowEmptyValues)
    m_ExpandableCategories = PropBag.ReadProperty("ExpandableCategories", m_def_ExpandableCategories)
    m_NameWidth = PropBag.ReadProperty("NameWidth", m_def_NameWidth)
    m_RequiresEnter = PropBag.ReadProperty("RequiresEnter", m_def_RequiresEnter)
    m_ShowCategories = PropBag.ReadProperty("ShowCategories", m_def_ShowCategories)
    m_ShowToolTips = PropBag.ReadProperty("ShowToolTips", m_def_ShowToolTips)
    m_CatBackColor = PropBag.ReadProperty("CatBackColor", m_def_CatBackColor)
    m_CatForeColor = PropBag.ReadProperty("CatForeColor", m_def_CatForeColor)
    m_SelBackColor = PropBag.ReadProperty("SelBackColor", m_def_SelBackColor)
    m_SelForeColor = PropBag.ReadProperty("SelForeColor", m_def_SelForeColor)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_GridColor = PropBag.ReadProperty("GridColor", m_def_GridColor)
    m_ExpandedImage = PropBag.ReadProperty("ExpandedImage", m_def_ExpandedImage)
    m_CollapsedImage = PropBag.ReadProperty("CollapsedImage", m_def_CollapsedImage)
'    m_ImageList = PropBag.ReadProperty("ImageList", m_def_ImageList)
    UserControl.Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    fGrid.Clear
    fGrid.Rows = 0
    fGrid.cols = ColCount
    m_Categories.Clear
    Grid_Config     ' config the grid
    Grid_Resize     ' resize the grid
    If (UserControl.Ambient.UserMode = False) Then
        m_bUserMode = False
        Set m_Properties = Nothing
        Set m_Properties = New Collection
        With m_Categories.Add("TPropertySheet")
            .Properties.Add "(Name)", UserControl.Ambient.DisplayName
            With .Properties.Add("Selected", "Value")
                .Selected = True
            End With
        End With
    Else
        m_bUserMode = True
    End If
End Sub

Private Sub UserControl_Show()
    Static s_bNotFirst As Boolean
    If Not (s_bNotFirst) Then
        ' set the parent of this resources
        SetParent lstBox.hwnd, Extender.Parent.hwnd
        SetParent monthView.hwnd, Extender.Parent.hwnd
        SetParent txtList.hwnd, Extender.Parent.hwnd
        ' stay on top
        StayOnTop lstBox.hwnd
        StayOnTop monthView.hwnd
        StayOnTop txtList.hwnd
        s_bNotFirst = True
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_font, Ambient.Font)
    Call PropBag.WriteProperty("AllowEmptyValues", m_AllowEmptyValues, m_def_AllowEmptyValues)
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
'    Call PropBag.WriteProperty("ImageList", m_ImageList, "")
    Call PropBag.WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
End Sub

Private Sub UserControl_Initialize()
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
    m_hStdIml = StdImages.hImageList
End Sub

Private Sub Grid_Initialize()
    ' set grid parameters for the sheet
    With fGrid
        '.Redraw = False
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
        .MergeCol(colStatus) = True
        '.Redraw = True
    End With
End Sub

Private Sub UserControl_Paint()
    On Error GoTo Err_UserControl_Paint
    Const constSource As String = m_constClassName & ".UserControl_Paint"
    
    Grid_Paint
    
    Exit Sub
Err_UserControl_Paint:
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

Private Sub UserControl_Resize()
    On Error GoTo Err_UserControl_Resize
    Const constSource As String = m_constClassName & ".UserControl_Resize"

'    UserControl_Paint
    Grid_Resize

    Exit Sub
Err_UserControl_Resize:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UserControl_Terminate()
    Set m_Categories = Nothing
    Set m_Properties = Nothing
End Sub

Private Sub Grid_Config()
    On Error GoTo Err_Grid_Config
    Const constSource As String = m_constClassName & ".Grid_Config"

    With fGrid
        '.Redraw = False
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
        Set .Font = m_font
        '.Redraw = True
    End With

    Exit Sub
Err_Grid_Config:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Grid_Resize()
    On Error GoTo Err_Grid_Resize
    Const constSource As String = m_constClassName & ".Grid_Resize"

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
        '.Redraw = False
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
            .ColWidth(colStatus) = 16 * Screen.TwipsPerPixelX
            wid = wid + .ColWidth(colStatus)
            cols = cols + 1
        Else
            .ColWidth(colStatus) = 0
        End If
        ' update col #1
        If m_hIml <> 0 Then
            .ColWidth(colPicture) = 0
        Else
            .ColWidth(colPicture) = 16 * Screen.TwipsPerPixelX
            wid = wid + .ColWidth(colPicture)
            cols = cols + 1
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
        ' update value picture column
        .ColWidth(colValuePicture) = 16 * Screen.TwipsPerPixelX
        wid = wid + .ColWidth(colValuePicture)
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
        '.Redraw = True
    End With

    Exit Sub
Err_Grid_Resize:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Grid_Clear()
    On Error GoTo Err_Grid_Clear
    Const constSource As String = m_constClassName & ".Grid_Clear"

    With fGrid
        .Clear
        .Rows = 0
        .cols = ColCount
    End With

    Exit Sub
Err_Grid_Clear:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Grid_Edit(KeyAscii As Integer, bFocus As Boolean)
    On Error GoTo Err_Grid_Edit
    Const constSource As String = m_constClassName & ".Grid_Edit"

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
    'fGrid.CellBackColor = vbWhite
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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Friend Sub AddNewCategory(objCat As TCategory)
    On Error GoTo Err_AddNewCategory
    Const constSource As String = m_constClassName & ".AddNewCategory"

    Dim strText As String
    Dim CurrRow As Integer
    Dim Ptr As Long
    Dim Index As Long
    Dim Cell As Integer
    
    ' stop flickering
    'fgrid.Redraw = False
    ' save row col position
    Cell = Cell_Save
    ' hide controls
    HideControls
    ' dehilite
'    DeHilite
    With fGrid
        ' add new category
        .AddItem "  " & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & ""
        ' the text for the cell is the category caption
        ' use the spaces to create distance from the picture
        strText = objCat.Caption
        ' set the row to update
        CurrRow = .Rows - 1
        ' get the pointer to the object
        Ptr = objCat.Handle
        ' get the catego
        Index = MakeDWord(objCat.Index, 0)
        ' row data will contain object pointer
        .RowData(CurrRow) = Ptr
        ' check for minus/plus picture to display
        If m_ShowCategories = True Then
            ' row is default
            .RowHeight(CurrRow) = DefaultHeight
            .ColWidth(colStatus) = 16 * Screen.TwipsPerPixelX
            SetState CurrRow, colStatus, objCat.Expanded
        Else
            .RowHeight(CurrRow) = 0
            .ColWidth(colStatus) = 0
        End If
        ' this row will be merged
        .MergeRow(CurrRow) = True
        ' set the current state for this category
        ' expanded/collapseed
        If objCat.Expanded Then
            objCat.Image = m_ExpandedImage
            Cell_DrawPicture CurrRow, colPicture, m_ExpandedImage
        Else
            objCat.Image = m_CollapsedImage
            Cell_DrawPicture CurrRow, colPicture, m_CollapsedImage
        End If
        ' write column #1
        Cell_Write CurrRow, colPicture, strText
        ' write column #2
        Cell_Write CurrRow, colName, strText
        ' write column #3
        Cell_Write CurrRow, colValuePicture, strText
        ' write column #4 (Value Column)
        Cell_Write CurrRow, colValue, strText
        ' write column #5 (sort column)
        Cell_Write CurrRow, colSort, Index
    End With
    ' restore cell position
    Cell_Restore Cell
    ' active drawing
    'fgrid.Redraw = True
    ' give way to windows
    DoEvents
    
    Exit Sub
Err_AddNewCategory:
    ' active drawing
    'fgrid.Redraw = True
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Friend Sub AddNewProperty(objProp As TProperty, bExpanded As Boolean)
    On Error GoTo Err_AddNewProperty
    Const constSource As String = m_constClassName & ".AddNewProperty"

    Dim strText As String
    Dim CurrRow As Integer
    Dim Index As Long
    Dim Ptr As Long
    Dim strDisplayText As String
    Dim Cell As Integer
    
    ' stop flickering
    'fgrid.Redraw = False
    ' save row col position
    Cell = Cell_Save
    ' hide controls
    HideControls
    ' dehilite
'    DeHilite
    ' update grid
    With fGrid
        strText = objProp.Caption
        ' set property's rowdata
        Ptr = objProp.Handle
        ' create unique index for this property
        Index = MakeDWord(objProp.Parent.Index, objProp.Index)
        ' add a new item
        .AddItem "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & 0
        strDisplayText = GetDisplayString(objProp)
        If strDisplayText = "" Then
            strDisplayText = " "
        End If
        ' save curr row
        CurrRow = .Rows - 1
        ' set handle to rowData to retrive the object later
        .RowData(CurrRow) = Ptr
        ' write column #1
        Cell_Write CurrRow, colPicture, strText
        ' write column #2
        Cell_Write CurrRow, colName, strText
        ' write column #3
        Cell_Write CurrRow, colValuePicture, strDisplayText
        ' write column #4 (formatted)
        Cell_Write CurrRow, colValue, strDisplayText
        ' write column #5
        Cell_Write CurrRow, colSort, Index
        ' merge this row
        .MergeRow(CurrRow) = True
        ' if ShowCategories is enabled check row height
        If m_ShowCategories = True Then
            ' if category is not expanded
            ' the cell height must be zero
            If bExpanded Then
                .RowHeight(CurrRow) = DefaultHeight
            Else
                .RowHeight(CurrRow) = 0    ' invisible
            End If
        Else
            .RowHeight(CurrRow) = DefaultHeight
        End If
        ' draw picture inside column #1
        If .RowHeight(CurrRow) <> 0 Then
            Cell_DrawPicture CurrRow, colPicture, objProp.Image
        End If
        ' add object pointer to properties collection
        m_Properties.Add Ptr, objProp.Caption
    End With
    ' restore cell position
    Cell_Restore Cell
    ' active drawing
    'fgrid.Redraw = True
    ' give way to windows
    DoEvents

    Exit Sub
Err_AddNewProperty:
    ' active drawing
    'fgrid.Redraw = True
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub fGrid_KeyDown(KeyCode As Integer, Shift As Integer)
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
    Const constSource As String = m_constClassName & ".fGrid_RowColChange"

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
                            UpdateProperty m_SelectedItem.ListValues(txtBox.Text).Value
                        End If
                    End If
                End If
            End If
        End If
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
'    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub fGrid_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Err_fGrid_MouseMove
    Const constSource As String = m_constClassName & ".fGrid_MouseMove"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub fGrid_DblClick()
    On Error GoTo Err_fGrid_DblClick
    Const constSource As String = m_constClassName & ".fGrid_DblClick"

    If m_bBrowseMode Or m_Categories.Count = 0 Then Exit Sub
    
    Dim Row As Integer
    
    ' get mouse row
    Row = fGrid.MouseRow
    ' if we are in a property cell
    If IsProperty(m_SelectedItem) Then
        ' check if we have to browse
        If IsBrowsable(m_SelectedItem) Then
            ' browse the property
            BrowseProperty
        Else
            ' get next avail row
            GetNextVisibleRowValue
        End If
    Else
        ' toggle category state expanded/collapsed
        ToggleCategoryState
    End If
    If txtBox.Visible Then
        txtBox.SetFocus
    End If
    Exit Sub
Err_fGrid_DblClick:
'    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub fGrid_Click()
    On Error GoTo Err_fGrid_Click
    Const constSource As String = m_constClassName & ".fGrid_Click"
    
    If m_bBrowseMode Or m_Categories.Count = 0 Then Exit Sub
    
    Dim Col As Integer
    
    ' get mouse coordinates in row/col
    Col = fGrid.MouseCol
    ' if its a category and the column is 0 then
    ' promote collapse/expand
    If m_SelectedItem Is Nothing Then
        Hilite fGrid.MouseRow
    End If
    If Not IsProperty(m_SelectedItem) And Col = 0 Then
        ToggleCategoryState
    End If

    Exit Sub
Err_fGrid_Click:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub fGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo Err_txtBox_KeyPress
    Const constSource As String = m_constClassName & ".txtBox_KeyPress"
    
    If m_Categories.Count = 0 Then Exit Sub
    
    Dim FindString As String
    Dim i As Integer
    Dim n As Integer
    
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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
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
    Const constSource As String = m_constClassName & ".txtBox_KeyPress"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub txtBox_LostFocus()
    If m_bBrowseMode = False Then
        UpdateProperty txtBox.Text
    End If
End Sub

Private Sub txtBox_DblClick()
    On Error GoTo Err_txtBox_DblClick
    Const constSource As String = m_constClassName & ".txtBox_DblClick"

    GetNextVisibleRowValue

    Exit Sub
Err_txtBox_DblClick:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub txtBox_Change()
    On Error GoTo Err_txtBox_Change
    Const constSource As String = m_constClassName & ".txtBox_Change"

    ' text has changed
    m_bDataChanged = True
    
    Exit Sub
Err_txtBox_Change:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub txtBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_txtBox_KeyDown
    Const constSource As String = m_constClassName & ".txtBox_KeyDown"

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
                   m_SelectedItem.ValueType = psCombo Then
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
                End If
            End If
    End Select

    Exit Sub
Err_txtBox_KeyDown:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
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
    Const constSource As String = m_constClassName & ".txtList_KeyDown"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub lstBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_lstBox_KeyDown
    Const constSource As String = m_constClassName & ".lstBox_KeyDown"

    Dim Index As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            UpdateProperty m_SelectedItem.ListValues(lstBox.ListIndex + 1).Value
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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub lstBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Err_lstBox_MouseUp
    Const constSource As String = m_constClassName & ".lstBox_MouseUp"
    If lstBox.ListIndex + 1 > 0 Then
        UpdateProperty m_SelectedItem.ListValues(lstBox.ListIndex + 1).Value, True
    End If
    Exit Sub
Err_lstBox_MouseUp:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub lstCheck_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Err_lstCheck_KeyDown
    Const constSource As String = m_constClassName & ".lstCheck_KeyDown"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub lstCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    UpdateCheckList
End Sub

Private Sub UpdateCheckList()
    On Error GoTo Err_UpdateCheckList
    Const constSource As String = m_constClassName & ".UpdateCheckList"
    Dim Value As String
    Dim i As Integer
    
    For i = 0 To lstCheck.ListCount - 1
        If lstCheck.Selected(i) Then
            If Value = "" Then
                Value = m_SelectedItem.ListValues(i + 1).Value
            Else
                Value = Value & " " & m_SelectedItem.ListValues(i + 1).Value
            End If
        End If
    Next
    
    UpdateProperty Value, True
    
    Exit Sub
Err_UpdateCheckList:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Function GetNextVisibleRowValue()
    On Error GoTo Err_GetNextVisibleRowValue
    Const constSource As String = m_constClassName & ".GetNextVisibleRowValue"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Function

Private Sub UpDown_DownClick()
    On Error GoTo Err_UpDown_DownClick
    Const constSource As String = m_constClassName & ".UpDown_DownClick"
    
    ' update with decreasing increment value
    UpdateUpDown -m_SelectedItem.UpDownIncrement
    
    Exit Sub
Err_UpDown_DownClick:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UpDown_UpClick()
    On Error GoTo Err_UpDown_UpClick
    Const constSource As String = m_constClassName & ".UpDown_UpClick"
    
    ' update with increasing increment
    UpdateUpDown m_SelectedItem.UpDownIncrement
    
    Exit Sub
Err_UpDown_UpClick:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
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
        If Not IsEmpty(MinValue) Then
            If Value < MinValue Then
                Value = MinValue
            End If
        End If
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
    Const constSource As String = m_constClassName & ".cmdBrowse_Click"

    BrowseProperty

    Exit Sub
Err_cmdBrowse_Click:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
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
    Const constSource As String = m_constClassName & ".UpdateProperty"
    
    ' If there is nothing to update then exit
    If (m_SelectedItem Is Nothing) Or _
       (m_bDataChanged = False And bForceUpdate = False) Then Exit Sub
    
    ' hide the controls activated by Grid_Edit()
    '    If m_SelectedItem.ValueType <> psDropDownCheckList Then
    '        'HideControls
    '    End If
    ' check for AllowEmptyValues
    If m_AllowEmptyValues = False And IsVarEmpty(NewValue) Then
        RaiseEvent EditError("Value for '" & m_SelectedItem.Caption & "' cannot be empty")
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
    RaiseEvent PropertyChanged(m_SelectedItem, NewValue, Cancel)
    ' permission denied get out
    If Cancel = True Then
        fGrid.SetFocus
        Exit Sub
    End If
    '    'StopFlicker hwnd
    Dim tmpValue As Variant
    ' data changed is false now
    m_bDataChanged = False
    ' check for a passed object here
    If IsObject(NewValue) Then
        Set m_SelectedItem.Value = NewValue
    Else
        tmpValue = ConvertValue(NewValue, m_SelectedItem.ValueType)
        If Not IsNull(tmpValue) Then
            If IsIncremental(m_SelectedItem) Then
                Dim MinValue
                Dim MaxValue
                m_SelectedItem.GetRange MinValue, MaxValue
                If Not IsEmpty(MinValue) Or Not IsEmpty(MaxValue) Then
                    If Not IsEmpty(MinValue) Then
                        If tmpValue < MinValue Then tmpValue = MinValue
                    End If
                    If Not IsEmpty(MaxValue) Then
                        If tmpValue > MaxValue Then
                            tmpValue = MaxValue
                        End If
                    End If
                End If
            End If
            m_SelectedItem.Value = tmpValue
        Else
            RaiseEvent EditError("Can't update. " & m_SelectedItem.Caption & " property has invalid data for its type")
        End If
    End If
    ' update textbox text value
    If m_SelectedItem.ValueType <> psDropDownCheckList Then
        UpdateTextBox m_SelectedItem
    End If
    '    'Release
    Exit Sub
Err_UpdateProperty:
    '    'Release
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub BrowseProperty()
    On Error GoTo Err_BrowseProperty
    Const constSource As String = m_constClassName & ".BrowseProperty"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub EditLongText()
    On Error GoTo Err_EditLongText
    Const constSource As String = m_constClassName & ".EditLongText"

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
    txtList.Height = 5 * TextHeight("A")
    txtList.Top = FixTopPos(txtList.Height)
    txtList.Text = strBuffer
    m_bDataChanged = False
    Set m_BrowseWnd = txtList
    txtList.Visible = True
    txtList.SetFocus

    Exit Sub
Err_EditLongText:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub EditCombo()
    On Error GoTo Err_EditCombo
    Const constSource As String = m_constClassName & ".EditCombo"

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
    If m_SelectedItem.ListValues.Count > 5 Then
        h = 5 * TextHeight("A")
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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub EditCheckList()
    On Error GoTo Err_EditCheckList
    Const constSource As String = m_constClassName & ".EditCheckList"

    Dim vArray As Variant
    Dim h As Integer
    Dim i As Integer
    Dim j As Integer
    Dim Value As String
    Dim t As Single
    
    lstCheck.Clear
    lstCheck.ZOrder
    SetControlFont lstCheck
    Value = m_SelectedItem.Value
    vArray = Split(Value, Chr(32))
    For i = 1 To m_SelectedItem.ListValues.Count
        lstCheck.AddItem m_SelectedItem.ListValues(i).Caption
        If Trim(m_SelectedItem.Value) <> "" Then
            If Not IsNull(vArray) Then
                For j = LBound(vArray) To UBound(vArray)
                    If m_SelectedItem.ListValues(i).Value = vArray(j) Then
                        lstCheck.Selected(lstCheck.NewIndex) = True
                    End If
                Next
            End If
        End If
    Next
    lstCheck.ListIndex = -1
    lstCheck.Width = rc.WindowWidth
    If m_SelectedItem.ListValues.Count > 5 Then
        h = 5 * TextHeight("A")
    Else
        h = (m_SelectedItem.ListValues.Count + 1) * TextHeight("A")
    End If
    lstCheck.Height = h
    t = FixTopPos(lstCheck.Height) - ((GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER)) * Screen.TwipsPerPixelY)
    '- ((GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYBORDER) * 3) * Screen.TwipsPerPixelY)
    lstCheck.Left = rc.WindowLeft '- (GetSystemMetrics(SM_CXBORDER) * 3 * Screen.TwipsPerPixelX)
    lstCheck.Top = t
    If BorderStyle = psBorderSingle Then
        lstCheck.Left = lstCheck.Left - (2 * Screen.TwipsPerPixelX)
        lstCheck.Top = lstCheck.Top - (2 * Screen.TwipsPerPixelY)
    End If
    Set m_BrowseWnd = lstCheck
    lstCheck.Visible = True
    lstCheck.SetFocus
    m_bDataChanged = False
    
    Exit Sub
Err_EditCheckList:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub UpdateTextBox(objProp As TProperty)
    On Error GoTo Err_UpdateTextBox
    Const constSource As String = m_constClassName & ".UpdateTextBox"

    Dim strDisplayStr As String
    '    'fgrid.Redraw = False
    '    txtBox.Visible = False
    ' get the display string for cell
    strDisplayStr = GetDisplayString(objProp)
    txtBox.Text = strDisplayStr
    m_bDataChanged = False
    ' update cell value
    '    If HasGraphicInterface(objProp) Then
    '        DrawGraphicInterface objProp
    '        strDisplayStr = Pad(strDisplayStr)
    '    End If
    '    If strDisplayStr = "" Then strDisplayStr = " "
    Cell_Write m_EditRow, colValuePicture, strDisplayStr
    Cell_Write m_EditRow, colValue, strDisplayStr
    'On Error Resume Next
    '    fGrid.SetFocus
    '    enter edit mode
    '    'fgrid.Redraw = True
    Grid_Edit 32, False
    SelectText
    Exit Sub
    
Err_UpdateTextBox:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub EditFont()
    On Error GoTo Err_EditFont
    Const constSource As String = m_constClassName & ".EditFont"

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
        .flags = CF_ScreenFonts
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
    Const constSource As String = m_constClassName & ".EditPicture"
    
    Dim strFileName As String
    Dim sTitle As String
    Dim sFilter As String
    Dim iFilterIndex As Integer
    Dim lFlags As Long
    Dim dlgCMD As New cCommonDialog
    
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
    ' call the event for user definition
    RaiseEvent BrowseForFile(m_SelectedItem, sTitle, sFilter, iFilterIndex, lFlags)
    With dlgCMD
        .Filter = sFilter
        .DialogTitle = sTitle
        .FilterIndex = iFilterIndex
        .Filename = strFileName
        .flags = lFlags
        .ShowOpen
        If Len(.Filename) > 0 Then
            UpdateProperty LoadPicture(.Filename), True
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
    Const constSource As String = m_constClassName & ".EditDate"

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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub monthView_DateClick(ByVal DateClicked As Date)
    UpdateProperty DateClicked, True
End Sub

Private Sub EditFile()
    On Error GoTo Err_EditFile
    Const constSource As String = m_constClassName & ".EditFile"

    Dim strFileName As String
    Dim sTitle As String
    Dim sFilter As String
    Dim iFilterIndex As Integer
    Dim lFlags As Long
    Dim dlgCMD As New cCommonDialog
    
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
    ' call the event for user defined vars
    RaiseEvent BrowseForFile( _
       m_SelectedItem, _
       sTitle, _
       sFilter, _
       iFilterIndex, _
       lFlags)
    ' update dialog properties
    With dlgCMD
        .Filter = sFilter
        .DialogTitle = sTitle
        .FilterIndex = iFilterIndex
        .Filename = strFileName
        .CancelError = True
        .flags = lFlags
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
    Const constSource As String = m_constClassName & ".EditFolder"

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
    Const constSource As String = m_constClassName & ".EditColor"
    
    Dim CurrColor As Long
    Dim dlgCMD As New cCommonDialog
    
    CurrColor = Val(m_SelectedItem.Value)
    With dlgCMD
        .DialogTitle = m_SelectedItem.Caption
        .CancelError = True
        .flags = CC_RGBInit
        .Color = CurrColor
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
    Const constSource As String = m_constClassName & ".ShowTextBox"
    
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
    Const constSource As String = m_constClassName & ".ShowBrowseButton"
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
    Const constSource As String = m_constClassName & ".ShowUpDown"

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
    Const constSource As String = m_constClassName & ".GetRowObject"

    Dim Ptr As Long
    
    If fGrid.Rows = 0 Then Exit Function
    Ptr = fGrid.RowData(Row)
    Set GetRowObject = ObjectFromPtr(Ptr)
    Exit Function

Err_GetRowObject:
'    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Function

Private Sub CollapseCategory(objCat As TCategory)
    On Error GoTo Err_CollapseCategory
    Const constSource As String = m_constClassName & ".CollapseCategory"

    Dim Cancel As Boolean
    Dim i As Integer
    Dim Row As Integer
    Dim Cell As Integer
    
    If ((objCat.Expanded = False) Or (m_ExpandableCategories = False)) Then Exit Sub
    Cancel = False
    RaiseEvent CategoryCollapsed(Cancel)
    If Cancel = True Then Exit Sub
    '    Row = FindGridRow(objCat)
    Row = m_SelectedRow
    'fgrid.Redraw = False
    Cell = Cell_Save
    With fGrid
        For i = 1 To objCat.Properties.Count
            .RowHeight(Row + i) = 0
            .Row = Row + 1
            .Col = colPicture
            'Set .CellPicture = Nothing
            .Col = colValuePicture
            'Set .CellPicture = Nothing
            '        Cell_ClearPicture Row + i, colPicture
            '        Cell_ClearPicture Row + i, colValuePicture
        Next
    End With
    SetState Row, colStatus, False
    objCat.Expanded = False
    ' set the apropriate icon
    objCat.Image = m_CollapsedImage
    Cell_DrawPicture Row, colPicture, m_CollapsedImage
    '    Grid_Resize
    Cell_Restore Cell
    'fgrid.Redraw = True
    
    Exit Sub
Err_CollapseCategory:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub ExpandCategory(objCat As TCategory)
    On Error GoTo Err_ExpandCategory
    Const constSource As String = m_constClassName & ".ExpandCategory"

    Dim Cancel As Boolean
    Dim i As Integer
    Dim Row As Integer
    Dim objProp As TProperty
    
    If ((objCat.Expanded = True) Or (m_ExpandableCategories = False)) Then Exit Sub
    Cancel = False
    RaiseEvent CategoryExpanded(Cancel)
    If Cancel = True Then Exit Sub
    '    'StopFlicker hwnd
    '    Row = FindGridRow(objCat)
    Row = m_SelectedRow
    'fgrid.Redraw = False
    For i = 1 To objCat.Properties.Count
        fGrid.RowHeight(Row + i) = -1 'DefaultHeight
        Set objProp = GetRowObject(Row + i)
        '        objProp.Selected = False
        Cell_DrawPicture Row + i, colPicture, objProp.Image
        If HasGraphicInterface(objProp) Then
            DrawGraphicInterface objProp, Row + i
        End If
    Next
    SetState Row, colStatus, True
    objCat.Expanded = True
    ' set the apropriate icon
    objCat.Image = m_ExpandedImage
    Cell_DrawPicture Row, colPicture, m_ExpandedImage
    Grid_Resize
    'fgrid.Redraw = True
    '    'Release
    
    Exit Sub
Err_ExpandCategory:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub SetState(Row As Integer, _
       Col As Integer, _
       bExpanded As Boolean)
       
    On Error GoTo Err_SetState
    Const constSource As String = m_constClassName & ".SetState"
    
    Dim Cell As Integer
    
    With fGrid
        '.Redraw = False
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
        '.Redraw = True
    End With
    
    Exit Sub
Err_SetState:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Function IsWindowLess(objProp As Object) As Boolean
    On Error GoTo Err_IsWindowLess
    Const constSource As String = m_constClassName & ".IsWindowLess"

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
    Const constSource As String = m_constClassName & ".GetNextVisibleRow"

    Dim Row As Integer
    Dim Obj As Object
    
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
    Const constSource As String = m_constClassName & ".GetPreviousVisibleRow"

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
    Const constSource As String = m_constClassName & ".DoSort"
    
    Dim Cell As Integer
    
    With fGrid
        '.Redraw = False
        Cell = Cell_Save
        .Row = RowStart
        .Col = Col
        .RowSel = RowEnd
        .ColSel = Col         ' fGrid.Cols - 1
        .Sort = SortMethod  ' 1 - Generic ascending.
        Cell_Restore Cell
        '.Redraw = True
    End With
end_sort:

    Exit Sub
Err_DoSort:
End Sub

Private Sub Hilite(Row As Integer)
    On Error GoTo Err_Hilite
    Const constSource As String = m_constClassName & ".Hilite"
    
    'StopFlicker hwnd
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
        '.Redraw = False
        .Row = Row                      ' set grid row
        .Col = colPicture                        ' set grid color col #1
        '        .ColSel = 2
        .CellBackColor = m_SelBackColor
        .CellForeColor = m_SelForeColor
        .Col = colName                        ' set grid color col #2
        .CellBackColor = m_SelBackColor
        .CellForeColor = m_SelForeColor
        .Col = colValuePicture                        ' set grid color col #2
        .CellBackColor = vbWhite
        .CellForeColor = vbBlack
        If IsProperty(m_SelectedItem) Then
            If HasGraphicInterface(m_SelectedItem) Then
                DrawGraphicInterface m_SelectedItem, Row
            End If
        End If
        ' save cell dimensions
        StoreCellPosition
        Cell_DrawPicture Row, colPicture, m_SelectedItem.Image
        '.Redraw = True
    End With
    'Release
    Exit Sub
Err_Hilite:
    'Release
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub DeHilite()
    On Error GoTo Err_DeHilite
    Const constSource As String = m_constClassName & ".DeHilite"

    Dim Obj As Object
    
    'StopFlicker hwnd
    ' finds the row associate with the selected object
    Set Obj = GetRowObject(m_SelectedRow)
    ' row not found then exit
    If Obj Is Nothing Then Exit Sub
    With fGrid
        '.Redraw = False
        .Row = m_SelectedRow            ' set grid row
        .Col = colPicture                        ' set grid col
        .ColSel = 4
        .CellBackColor = Obj.BackColor
        .CellForeColor = Obj.ForeColor
        Obj.Selected = False
        If HasGraphicInterface(m_SelectedItem) Then
            DrawGraphicInterface m_SelectedItem, m_SelectedRow
        End If
        Cell_DrawPicture m_SelectedRow, colPicture, Obj.Image
        '.Redraw = True
    End With
    'Release
    Exit Sub

Err_DeHilite:
    'Release
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Friend Function FindGridRow(Obj As Object) As Integer
    On Error GoTo Err_FindGridRow
    Const constSource As String = m_constClassName & ".FindGridRow"

    Dim i As Integer
    Dim Ptr As Long
    
    Ptr = Obj.Handle
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

Private Sub Cell_Write( _
       ByVal Row As Integer, _
       ByVal Col As Integer, _
       ByVal strText As String)
    
    Dim Obj As Object
    Dim Cell As Integer
    
    Set Obj = GetRowObject(Row)
    If Obj Is Nothing Then Exit Sub
    strText = Trim(strText)
    'fgrid.Redraw = False
    Cell = Cell_Save
    If TypeOf Obj Is TCategory Then
        Cell_WriteCategory Obj, Row, Col, strText
    Else
        Cell_WriteProperty Obj, Row, Col, strText
    End If
    Cell_Restore Cell
    'fgrid.Redraw = True
End Sub

Private Sub Cell_WriteCategory( _
       CatObj As TCategory, _
       ByVal Row As Integer, _
       ByVal Col As Integer, _
       ByVal strText As String)
                      
    On Error GoTo Err_Cell_Write
    Const constSource As String = m_constClassName & ".Cell_Write"

    Dim tmpBackColor As OLE_COLOR
    Dim tmpForecolor As OLE_COLOR
    Dim ObjFont As StdFont
    
    ' configure back color
    If CatObj.Selected Then
        tmpBackColor = m_SelBackColor
        tmpForecolor = m_SelForeColor
    Else
        If CatObj.BackColor = CLR_INVALID Then
            tmpBackColor = m_CatBackColor
        Else
            tmpBackColor = CatObj.BackColor
        End If
        If CatObj.ForeColor = CLR_INVALID Then
            tmpForecolor = m_CatForeColor
        Else
            tmpForecolor = CatObj.ForeColor
        End If
    End If
    Set ObjFont = m_CatFont
    strText = Pad(strText)
    ' write the cell
    With fGrid
        .Row = Row
        .Col = Col
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForecolor
        .CellAlignment = flexAlignLeftCenter
        If Not ObjFont Is Nothing Then
            .CellFontName = ObjFont.Name
            .CellFontBold = ObjFont.Bold
            .CellFontItalic = ObjFont.Italic
            .CellFontStrikeThrough = ObjFont.Strikethrough
            .CellFontUnderline = ObjFont.Underline
            .CellFontSize = ObjFont.Size
        End If
        .Text = strText
    End With
    
    Exit Sub
Err_Cell_Write:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Cell_WriteProperty( _
       PropObj As TProperty, _
       ByVal Row As Integer, _
       ByVal Col As Integer, _
       ByVal strText As String)
                      
    On Error GoTo Err_Cell_Write
    Const constSource As String = m_constClassName & ".Cell_Write"

    Dim tmpBackColor As OLE_COLOR
    Dim tmpForecolor As OLE_COLOR
    
    ' configure back color
    If PropObj.Selected Then
        If Col < colValuePicture Then
            tmpBackColor = m_SelBackColor
            tmpForecolor = m_SelForeColor
        Else
            tmpBackColor = vbWhite
            tmpForecolor = vbBlack
        End If
    Else
        If PropObj.BackColor = CLR_INVALID Then
            tmpBackColor = m_BackColor
        Else
            tmpBackColor = PropObj.BackColor
        End If
        If PropObj.ForeColor = CLR_INVALID Then
            tmpForecolor = m_ForeColor
        Else
            tmpForecolor = PropObj.ForeColor
        End If
    End If
    ' check what col to write
    If Col = colValue Or Col = colValuePicture Then
        If HasGraphicInterface(PropObj) Then
            DrawGraphicInterface PropObj, Row
            strText = Space(m_lPadding) & strText
        Else
            ' clear picture
            fGrid.Col = colValuePicture
            fGrid.Row = Row
'            If Not fGrid.CellPicture Is Nothing Then
'                Set fGrid.CellPicture = Nothing
'            End If
            If strText = "" Then
                strText = Space(m_lPadding) & strText
            End If
        End If
    ElseIf Col > colStatus And Col < colSort Then
        strText = Pad(strText)
    End If
    ' write the cell
    With fGrid
        'fGrid.cols = ColCount
        .Row = Row
        .Col = Col
        .CellBackColor = tmpBackColor
        .CellForeColor = tmpForecolor
        .CellAlignment = flexAlignLeftCenter
        .Text = strText
    End With
    
    Exit Sub
Err_Cell_Write:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Cell_WriteValue( _
       ByVal objProp As TProperty, _
       ByVal Row As Integer)

    On Error GoTo Err_Cell_WriteValue
    Const constSource As String = m_constClassName & ".Cell_WriteValue"

    Dim strDisplayStr As String
   
    With fGrid
        strDisplayStr = GetDisplayString(objProp)
        Cell_Write Row, colValuePicture, strDisplayStr
        Cell_Write Row, colValue, strDisplayStr
    End With

    Exit Sub
Err_Cell_WriteValue:
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Cell_DrawPicture( _
        Row As Integer, _
        Col As Integer, _
        Image As Variant)
    
    On Error GoTo Err_Cell_DrawPicture
    Const constSource As String = m_constClassName & ".Cell_DrawPicture"
    
    Dim BkColor As OLE_COLOR
    Dim Obj As Object
    
    Set Obj = GetRowObject(Row)
    If Obj Is Nothing Then Exit Sub
    If Obj.Selected Then
        BkColor = m_SelBackColor
    Else
        BkColor = Obj.BackColor
    End If
    Cell_DrawPictureEx Row, Col, Image, m_hIml, BkColor
    Exit Sub

Err_Cell_DrawPicture:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub Cell_ClearPicture( _
    Row As Integer, _
    Col As Integer)
    
    On Error GoTo Err_Cell_ClearPicture
    Const constSource As String = m_constClassName & ".Cell_ClearPicture"
    
    Dim Obj As Object
    Dim tmpBackColor As OLE_COLOR
    Dim Cell As Integer
        
    Set Obj = GetRowObject(Row)
    If Obj Is Nothing Then Exit Sub
    With fGrid
        '.Redraw = False
        Cell = Cell_Save
        .Row = Row
        .Col = Col
        Set .CellPicture = Nothing
        Cell_Restore Cell
        '.Redraw = True
    End With

    Exit Sub
Err_Cell_ClearPicture:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Function IsProperty(ByVal Obj As Object) As Boolean
    IsProperty = TypeName(Obj) = "TProperty"
End Function

'Friend Sub ValueChanged(Prop As TProperty)
'    On Error GoTo Err_ValueChanged
'    Const constSource As String = m_constClassName & ".ValueChanged"
'
'    Dim Row As Integer
'    Dim Cell As Integer
'
'    ' disable drawing
'    'fgrid.Redraw = False
'    ' save cell pos
'    Cell = Cell_Save
'    ' get object row
'    Row = FindGridRow(Prop)
'    ' invalid row then exit
'    If Row = -1 Then Exit Sub
'    ' check for the value type
'    If (Prop.ValueType <> psDropDownCheckList) Then
'        HideBrowseWnd
'        txtBox.Visible = False
'        If UpDown.Visible = False And cmdBrowse.Visible = False Then
'            HideControls
'        End If
'    End If
'    Cell_Write Row, colName, Prop.Caption
'    ' write the edit cell
'    Cell_WriteValue Prop, Row
'    ' restore cell properties
'    Cell_Restore Cell
'    ' enable drawing
'    'fgrid.Redraw = True
'
'    Exit Sub
'Err_ValueChanged:
'    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
'End Sub
'
'Friend Sub CaptionChanged(Prop As TProperty)
'    On Error GoTo Err_ValueChanged
'    Const constSource As String = m_constClassName & ".ValueChanged"
'
'    Dim Row As Integer
'    Dim Cell As Integer
'
'    ' disable drawing
'    'fgrid.Redraw = False
'    ' save cell pos
'    Cell = Cell_Save
'    ' get object row
'    Row = FindGridRow(Prop)
'    ' invalid row then exit
'    If Row = -1 Then Exit Sub
'    ' check for the value type
'    If (Prop.ValueType <> psDropDownCheckList) Then
'        HideBrowseWnd
'        txtBox.Visible = False
'        If UpDown.Visible = False And cmdBrowse.Visible = False Then
'            HideControls
'        End If
'    End If
'    ' write the edit cell
'    Cell_Write Row, colName, Prop.Caption
'    ' restore cell properties
'    Cell_Restore Cell
'    ' enable drawing
'    'fgrid.Redraw = True
'
'    Exit Sub
'Err_ValueChanged:
'    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
'End Sub

'Private Sub Grid_ChangeBackColor(NewColor As OLE_COLOR, bCategory As Boolean)
'    Dim Row As Integer
'    Dim Obj As Object
'    Dim bChange As Boolean
'    Dim Cell As Integer
'
'    'StopFlicker hwnd
'    'fgrid.Redraw = False
'    Cell = Cell_Save
'    For Row = 0 To fGrid.Rows - 1
'        If bCategory Then
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) = 0
'        Else
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) > 0
'        End If
'        If bChange Then
'            Set Obj = GetRowObject(Row)
'            If Obj.BackColor <> m_BackColor Then
'                Obj.BackColor = NewColor
'                fGrid.Row = Row
'                fGrid.Col = colPicture ' 1
'                fGrid.ColSel = colValuePicture ' 3
'                If Obj.Selected = False Then
'                    fGrid.CellBackColor = NewColor
'                Else
'                    fGrid.CellBackColor = m_SelBackColor
'                End If
'                Cell_DrawPicture Row, colPicture, Obj.Image
'            End If
'        End If
'    Next
'    Cell_Restore Cell
'    'fgrid.Redraw = True
'    'Release
'End Sub
'
'Private Sub Grid_ChangeForeColor(NewColor As OLE_COLOR, bCategory As Boolean)
'    Dim Row As Integer
'    Dim Obj As Object
'    Dim bChange As Boolean
'    Dim Cell As Integer
'
'    'StopFlicker hwnd
'    DeHilite
'    Cell = Cell_Save
'    'fgrid.Redraw = False
'    For Row = 0 To fGrid.Rows - 1
'        If bCategory Then
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) = 0
'        Else
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) > 0
'        End If
'        If bChange Then
'            Set Obj = GetRowObject(Row)
'            Obj.ForeColor = NewColor
'            fGrid.Row = Row
'            fGrid.Col = colPicture
'            fGrid.ColSel = colValuePicture
'            If Obj.Selected = False Then
'                fGrid.CellForeColor = NewColor
'            Else
'                fGrid.CellForeColor = m_SelForeColor
'            End If
'            Cell_DrawPicture Row, colPicture, Obj.Image
'        End If
'    Next
'    Cell_Restore Cell
'    'fgrid.Redraw = True
'    'Release
'End Sub
'
'Private Sub Grid_ChangeSelForeColor(NewColor As OLE_COLOR)
'    Dim Row As Integer
'    Dim Obj As Object
'    Dim bChange As Boolean
'    Dim Cell As Integer
'
'    DeHilite
'    Cell = Cell_Save
'    'fgrid.Redraw = False
'    For Row = 0 To fGrid.Rows - 1
'        Set Obj = GetRowObject(Row)
'        If Obj.Selected Then
'            fGrid.Row = Row
'            fGrid.Col = colPicture
'            fGrid.ColSel = colPicture
'            fGrid.CellForeColor = NewColor
'            Cell_DrawPicture Row, colPicture, Obj.Image
'            Exit For
'        End If
'    Next
'    Cell_Restore Cell
'    'fgrid.Redraw = True
'End Sub
'
'Private Sub Grid_ChangeSelBackColor(NewColor As OLE_COLOR)
'    Dim Row As Integer
'    Dim Obj As Object
'    Dim Cell As Integer
'
'    '    'StopFlicker hwnd
'    DeHilite
'    For Row = 0 To fGrid.Rows - 1
'        Set Obj = GetRowObject(Row)
'        If Obj.Selected Then
'            Cell = Cell_Save
'            'fgrid.Redraw = False
'            fGrid.Row = Row
'            fGrid.Col = colPicture
'            fGrid.ColSel = colName
'            fGrid.CellBackColor = NewColor
'            Cell_DrawPicture Row, colPicture, Obj.Image
'            Cell_Restore Cell
'            'fgrid.Redraw = True
'            Exit For
'        End If
'    Next
'    '    'Release
'End Sub
'
'Private Sub Grid_ChangeFont(bCategory As Boolean)
'    Dim Row As Integer
'    Dim bChange As Boolean
'    Dim Cell As Integer
'
'    'StopFlicker hwnd
'    Cell = Cell_Save
''    Set fGrid.Font = m_Font
'    'fgrid.Redraw = False
'    For Row = 0 To fGrid.Rows - 1
'        If bCategory Then
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) = 0
'        Else
'            bChange = LoWord(fGrid.TextMatrix(Row, colSort)) > 0
'        End If
'        If bChange Then
'            Debug.Print StrFromFont(fGrid.Font)
'            Cell_Write Row, colPicture, fGrid.TextMatrix(Row, colPicture)
'            Cell_Write Row, colName, fGrid.TextMatrix(Row, colName)
'            Cell_Write Row, colValuePicture, fGrid.TextMatrix(Row, colValuePicture)
'            Cell_Write Row, colValue, fGrid.TextMatrix(Row, colValue)
'        End If
'    Next
'    Cell_Restore Cell
'    'fgrid.Redraw = True
'    'Release
'End Sub
'
'Private Sub Grid_ChangeCategoryImage(NewImage, bExpanded As Boolean)
'    Dim Row As Integer
'    Dim Obj As TCategory
'    Dim Cell As Integer
'
'    Cell = Cell_Save
'    'fgrid.Redraw = False
'    For Row = 0 To fGrid.Rows - 1
'        If LoWord(fGrid.TextMatrix(Row, colSort)) = 0 Then
'            Set Obj = GetRowObject(Row)
'            If Obj.Expanded = bExpanded Then
'                Obj.Image = NewImage
'                Cell_DrawPicture Row, colPicture, Obj.Image
'            End If
'        End If
'    Next
'    Cell_Restore Cell
'    'fgrid.Redraw = True
'End Sub
'
'Private Sub Grid_ChangeImageList()
'    Dim Row As Integer
'    Dim Obj As Object
'    Dim Cell As Integer
'
'    ' save cell position
'    Cell = Cell_Save
'    ' deactivate drawing
'    'fgrid.Redraw = False
'    ' loop the rows to draw picture
'    For Row = 0 To fGrid.Rows - 1
'        Set Obj = GetRowObject(Row)
'        Cell_DrawPicture Row, colPicture, Obj.Image
'    Next
'    ' restore cell pos
'    Cell_Restore Cell
'    ' enable drawing
'    'fgrid.Redraw = True
'End Sub
'
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
    Const constSource As String = m_constClassName & ".GetDefaultDisplayString"

    Dim strDisplayStr As String
    Dim i As Integer
    Dim strTemp As String
    
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
            For i = 1 To objProp.ListValues.Count
                If objProp.Value = objProp.ListValues(i).Value Then
                    strDisplayStr = objProp.ListValues(i).Caption
                    GoTo Exit_GetDefaultDisplayString
                End If
            Next
            strDisplayStr = objProp.Value
        Case psDropDownList
            For i = 1 To objProp.ListValues.Count
                If objProp.Value = objProp.ListValues(i).Value Then
                    strDisplayStr = objProp.ListValues(i).Caption
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
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
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
        .Col = colValuePicture
        ' get cell rect
        rc.Left = .CellLeft
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
        rc.InterfaceLeft = .CellLeft + edgeX + TextWidth(Space(m_lPadding)) '(16 * Screen.TwipsPerPixelX)
    End With
End Sub

Private Sub Grid_ShowCategories()
    On Error GoTo Err_ShowCategories
    Dim Row As Integer
    Dim Cell As Integer
    Dim objCat As TCategory
    Dim i As Integer
    Dim Prop As TProperty
    Dim RowStart As Integer
    Dim RowSel As Integer
    Dim objTemp As Object
    
    ' stop screen flickering
    'StopFlicker hwnd
    ' get the object related with this row
    Set objTemp = GetRowObject(m_SelectedRow)
    ' hide all the controls
    HideControls
    ' save cell position
    Cell = Cell_Save
    ' stop redrawing
    'fgrid.Redraw = False
    ' if it is to disable categories...
    If m_ShowCategories = False Then
        ' set column #0 width to 0
        fGrid.ColWidth(colStatus) = 0
        If fGrid.Rows > 0 Then
            For Row = 0 To fGrid.Rows - 1
                ' check it is a category
                If LoWord(Val(fGrid.TextMatrix(Row, colSort))) = 0 Then
                    ' clear minus/plus picture
                    'Cell_ClearPicture Row, colStatus
                    ' clear category picture
                    'Cell_ClearPicture Row, ColPicture
                    ' row is invisible
                    fGrid.RowHeight(Row) = 0
                Else
                    ' force all property rows to be displayed
                    fGrid.RowHeight(Row) = DefaultHeight
                End If
            Next
            ' sort entire row count
            DoSort 0, fGrid.Rows - 1, colName, 1
        End If
    Else
        ' start with the first row
        Row = 0
        ' restore column #0 width
        fGrid.ColWidth(colStatus) = -1
        ' sort by colum #4 (index column)
        DoSort 0, fGrid.Rows - 1, colSort, 1
        ' loop the row to make categories visible
        Do While Row < fGrid.Rows
            'On Error Resume Next
            ' get the category object
            Set objCat = GetRowObject(Row)
            If objCat Is Nothing Then
                Row = Row + 1
                'Exit Sub
            Else
                ' set the state
                SetState Row, colStatus, objCat.Expanded
                ' turn row visible again
                fGrid.RowHeight(Row) = DefaultHeight
                ' draw category picture
                Cell_DrawPicture Row, colPicture, objCat.Image
                Row = Row + 1
                If Row >= fGrid.Rows Then Exit Do
                RowStart = Row
                Do While LoWord(fGrid.TextMatrix(Row, colSort)) > 0
                    ' if the category is not expanded then
                    ' set all property to invisible
                    If objCat.Expanded = False Then
                        fGrid.RowHeight(Row) = 0
                    Else
                        ' get the property
                        Set Prop = GetRowObject(Row)
                        ' draw category picture
                        Cell_DrawPicture Row, colPicture, Prop.Image
                        ' if this property has graphic interface
                        ' draw the interface
                        If HasGraphicInterface(Prop) Then
                            DrawGraphicInterface Prop, Row
                        End If
                    End If
                    Row = Row + 1
                    If Row >= fGrid.Rows Then Exit Do
                Loop
                ' sort by colum #2 (names column) within this category
                If Row - 1 > RowStart Then
                    DoSort RowStart, Row - 1, colName, 1
                End If
            End If
        Loop
    End If
    ' select the object row
    If Not objTemp Is Nothing Then
        m_SelectedRow = FindGridRow(objTemp)
    Else
        If fGrid.Rows > 0 Then
            m_SelectedRow = 0
        End If
    End If
    ' restore position
    Cell_Restore Cell
    ' resize the grid
    Grid_Resize
    ' activate drawing
    'fgrid.Redraw = True
Err_ShowCategories:
    ' 'Release painting
    'Release
End Sub

Private Sub SetControlFont(Ctl As Control)
    Ctl.FontName = m_font.Name
    Ctl.FontSize = m_font.Size
End Sub

Private Function FixTopPos(lHeight) As Long
    lHeight = CLng(lHeight)
    If rc.WindowTop + lHeight + 300 > UserControl.Extender.Parent.ScaleHeight Then
        FixTopPos = rc.WindowTop - (rc.Height + lHeight)
    Else
        FixTopPos = rc.WindowTop
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
        BkColor = objProp.BackColor
    End If
    If fGrid.ColWidth(colValuePicture) = 0 Then
        fGrid.ColWidth(colValuePicture) = -1
    End If
    Cell_DrawPictureEx Row, colValuePicture, Image, StdImages.hImageList, BkColor
End Sub

Private Sub DrawColorBox(Row As Integer, _
       objProp As TProperty)
    Dim Image As String
    Dim BkColor As OLE_COLOR
    
    Image = "frame"
    BkColor = objProp.Value
    Cell_DrawPictureEx Row, colValuePicture, StdImages.ListImages(Image).Index, StdImages.hImageList, BkColor
End Sub

Private Sub Cell_DrawPictureEx( _
       ByVal Row As Integer, _
       ByVal Col As Integer, _
       ByVal Image As Long, _
       hIml As Long, _
       Optional BkColor As OLE_COLOR = 0)
    
    On Error GoTo Err_Cell_DrawPictureEx
    Const constSource As String = m_constClassName & ".Cell_DrawPictureEx"
    
    Dim varX As Variant
    Dim Obj As Object
    Dim hIcon As Long
    Dim lhDC As Long
    Dim lX As Long
    Dim lY As Long
    Dim rc As RECT
    Dim rw As RECT
    
    If fGrid.RowHeight(Row) = 0 Or hIml = 0 Then Exit Sub
'    On Error Resume Next
'    varX = ImageList_GetIcon(hIml, Image, 0)
'    If Err.Number <> 0 Then Exit Sub
    Set Obj = GetRowObject(Row)
    If Obj Is Nothing Then Exit Sub
    ' First get the Desktop DC:
    lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
    ' Set the draw mode to XOR:
    SetROP2 lhDC, R2_NOTXORPEN
    GetClientToScreen Extender.Parent.hwnd, rw
    GetClientToScreen Extender.hwnd, rc
    With fGrid
        ''.Redraw = False
        .Row = Row
        .Col = Col
        lX = rc.Left + (.CellLeft / Screen.TwipsPerPixelX)
        lY = rc.Top + (.CellTop / Screen.TwipsPerPixelY)
        ImageList_Draw m_hIml, Image, lhDC, lX, lY, ILD_TRANSPARENT
'        DrawImage Image, lhDC, lX, lY
        ''.Redraw = True
    End With
    DeleteDC lhDC
    
    Exit Sub
Err_Cell_DrawPictureEx:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Private Sub RecalcPadding()
    Dim Twips As Long
    Dim w As Single
    Dim s As Single
    Dim str As String
    
    If Not m_font Is Nothing Then
        Twips = 16 * Screen.TwipsPerPixelX
        Set UserControl.Font = m_font
        Do
            str = str & " "
            s = TextWidth(str)
            w = w + 1
        Loop Until s > Twips
        m_lPadding = w '- 1
    End If
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
    Const constSource As String = m_constClassName & ".LoadFromFile"

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
    AllowEmptyValues = ReadProperty(Col("AllowEmptyValues"), m_def_AllowEmptyValues)
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
    Err.Raise Description:="Unexpected Error: " & Err.Description, Number:=Err.Number, Source:=constSource
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
    Const constSource As String = m_constClassName & ".SaveFile"

    Dim i As Integer
    Dim j As Integer
    Dim hFile As Integer
    
    'StopFlicker hwnd
    m_strText = "[" & Section & "]" & vbCrLf
    Call WriteProperty("Font", StrFromFont(m_font))
    Call WriteProperty("CatFont", StrFromFont(m_CatFont))
    Call WriteProperty("AllowEmptyValues", m_AllowEmptyValues)
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
    Err.Raise Description:=Err.Description, _
       Number:=Err.Number, _
       Source:=constSource
End Sub

Private Sub WriteProperty(ByVal Prop As String, ByVal Value As Variant)
    m_strText = m_strText & Prop & "=" & Value & vbCrLf
End Sub

Private Function DefaultHeight() As Long
    ' return height in pixels
    DefaultHeight = -1 '17 * Screen.TwipsPerPixelY
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

Friend Sub Cell_ChangeColor(Obj As Object)
    Dim Row As Integer
    
    Row = FindGridRow(Obj)
    If Row = -1 Then Exit Sub
    With fGrid
        '.Redraw = False
        .Row = Row            ' set grid row
        .Col = colPicture                        ' set grid col
        .ColSel = 4
        .CellBackColor = Obj.BackColor
        .CellForeColor = Obj.ForeColor
        .Col = colName                        ' set grid col
        .CellBackColor = Obj.BackColor
        .CellForeColor = Obj.ForeColor
        Cell_DrawPicture Row, colPicture, Obj.Image
        If HasGraphicInterface(Obj) Then
            DrawGraphicInterface Obj, Row
        End If
        '.Redraw = True
    End With
End Sub

Private Function Pad(ByVal Text As String) As String
    If m_hIml = 0 Then
        Pad = Text
    Else
        Pad = Space(m_lPadding) & Text
    End If
End Function

Private Function AvgTextWidth() As Single
    Dim avgWidth As Single
    ' Get the average character width of the current list box Font
    ' (in pixels) using the form's TextWidth width method.
    avgWidth = UserControl.TextWidth("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")
    avgWidth = avgWidth / 52
    ' Set the white space you want between columns.
    AvgTextWidth = avgWidth
End Function

Private Function EvaluateTextWidth(ByVal s As String) As Single
    Dim avgWidth As Single
    
    avgWidth = AvgTextWidth
    EvaluateTextWidth = Len(Trim(s)) * avgWidth
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
    Const constSource As String = m_constClassName & ".Grid_Paint"
    
    Dim Cat As Integer
    Dim Prop As Integer
    Dim RowStart As Integer
    Dim RowSel As Integer
    Dim Cell As Integer
    Dim Row As Integer
    Dim Obj As Object
    
    If m_bDirty = False Then Exit Sub
    m_bDirty = False
    RecalcPadding
    ' check for Categories. If no categories
    ' is found then clear grid and exit
    If m_Categories.Count = 0 Then
        Grid_Clear
        Exit Sub
    End If
    ' hide all visible control
    HideControls
    ' draw grid cells
    With fGrid
        ' to avoid flickering
        '.Redraw = False
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
        Set .Font = m_font
        ' loop the categories and properties configuring each one
        For Row = 0 To .Rows - 1
            Cell_Write Row, colPicture, .TextMatrix(Row, colPicture)
            Cell_Write Row, colName, .TextMatrix(Row, colName)
            Cell_Write Row, colValuePicture, .TextMatrix(Row, colValuePicture)
            Cell_Write Row, colValue, .TextMatrix(Row, colValue)
        Next
        ' restore cell position
        Cell_Restore Cell
        '.Redraw = True
    End With

    Exit Sub
Err_Grid_Paint:
    Err.Raise Description:=Err.Description, Number:=Err.Number, Source:=constSource
End Sub

Friend Sub TriggerEvent(ByVal RaisedEvent As String, ParamArray aParams())
    Select Case RaisedEvent
        Case "CaptionChanged"
    End Select
End Sub

Public Sub DrawImage( _
        ByVal vKey As Variant, _
        ByVal hdc As Long, _
        ByVal xPixels As Integer, _
        ByVal yPixels As Integer, _
        Optional ByVal bSelected = False, _
        Optional ByVal bCut = False, _
        Optional ByVal bDisabled = False, _
        Optional ByVal oCutDitherColour As OLE_COLOR = vbWindowBackground, _
        Optional ByVal hExternalIml As Long = 0 _
    )
    Dim hIcon As Long
    Dim lFlags As Long
    Dim lhIml As Long
    Dim lColor As Long
    Dim iImgIndex As Long

   ' Draw the image at 1 based index or key supplied in vKey.
   ' on the hDC at xPixels,yPixels with the supplied options.
   ' You can even draw an ImageList from another ImageList control
   ' if you supply the handle to hExternalIml with this function.
   
   iImgIndex = Val(vKey)
   
   If (iImgIndex > -1) Then
      If (hExternalIml <> 0) Then
          lhIml = 0 'hExternalIml
      Else
          lhIml = m_hIml
      End If
      lFlags = ILD_TRANSPARENT
      If (bSelected) Or (bCut) Then
          lFlags = lFlags Or ILD_SELECTED
      End If
      If (bCut) Then
        ' Draw dithered:
        lColor = TranslateColor(oCutDitherColour)
        If (lColor = -1) Then lColor = GetSysColor(COLOR_WINDOW)
        ImageList_DrawEx _
              lhIml, _
              iImgIndex, _
              hdc, _
              xPixels, yPixels, 0, 0, _
              CLR_NONE, lColor, _
              lFlags
      ElseIf (bDisabled) Then
'        ' extract a copy of the icon:
'        hIcon = ImageList_GetIcon(m_hIml, iImgIndex, 0)
'        ' Draw it disabled at x,y:
'        DrawState hdc, 0, 0, hIcon, 0, xPixels, yPixels, m_lIconSizeX, m_lIconSizeY, DST_ICON Or DSS_DISABLED
'        ' Clear up the icon:
'        DestroyIcon hIcon
      Else
        ' Standard draw:
        ImageList_Draw _
            m_hIml, _
            iImgIndex, _
            hdc, _
            xPixels, _
            yPixels, _
            lFlags
      End If
   End If
End Sub

Private Sub GetClientToScreen(hWndA As Long, rc As RECT)
    Dim tp As POINTAPI
    
    ' Get the client rectangle of the window in screen coordinates:
    GetClientRect hWndA, rc
    tp.x = rc.Left
    tp.y = rc.Top
    ClientToScreen hWndA, tp
    rc.Left = tp.x
    rc.Top = tp.y
    tp.x = rc.Right
    tp.y = rc.Bottom
    ClientToScreen hWndA, tp
    rc.Right = tp.x
    rc.Bottom = tp.y
End Sub
'-- end code
