VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TProperty"
Attribute VB_Ext_KEY = "Member0" ,"TProperty"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
' 类模块       : TProperties.cls
' 作      者   : Marclei V Silva (MVS)
' 程序员       : Marclei V Silva (MVS) [Spnorte Consultoria de Inform醫ica]
' 编 写 日 期  : 06/16/2000 -- 08:27:55
' 输  入       : N/A
' 输  出       : N/A
' 描   述      : Collection of properties
'              :
'              :
'              :
' *******************************************************
Option Explicit
Option Compare Text

' local variable to hold collection
Private mCol As Collection
Private m_ParentPtr As Long

Friend Sub Init(Object As TCategory)
    m_ParentPtr = PtrFromObject(Object)
End Sub

Public Property Get Parent() As TCategory
    Set Parent = ObjectFromPtr(m_ParentPtr)
End Property

Public Function Add( _
    Key As String, _
    Caption As String, _
    Optional Value As Variant, _
    Optional ValueType As psPropertyType = psNone, _
    Optional ReadOnly As Boolean, _
    Optional Image As Variant = 0, _
    Optional TooltipText As String _
    ) As TProperty
    
    Dim objNewMember As TProperty
    Dim Index As Integer
    Dim i As Integer
    
    ' create a new object
    Set objNewMember = New TProperty
    ' check for no given value type
    If ValueType = psNone Then
        ValueType = GetValueType(Value)
    End If
    ' set the properties passed into the method
    objNewMember.Caption = Caption
    objNewMember.Image = Image
    objNewMember.ReadOnly = ReadOnly
    objNewMember.TooltipText = TooltipText
    objNewMember.ValueType = ValueType
    If IsMissing(Value) Then
        Value = GetDefaultValue(ValueType)
    Else
        If IsObject(Value) Then
            Set objNewMember.Value = Value
        Else
            objNewMember.Value = Value
        End If
    End If
    ' if it is a boolean value then add items
    ' to ListValues collection
    If ValueType = psBoolean Then
        objNewMember.ListValues.Add False, "False", "False"
        objNewMember.ListValues.Add True, "True", "True"
    ElseIf ValueType = psTime Then
        objNewMember.UpDownIncrement = 1
        objNewMember.Format = "hh:mm:ss"
    ElseIf ValueType = psCurrency Then
        ' value type is currency then add a format
        objNewMember.Format = "$ #,##0.00"
    End If
    ' set the default properties
    objNewMember.BackColor = CLR_INVALID
    objNewMember.ForeColor = CLR_INVALID
    objNewMember.MaskColor = CLR_INVALID
    objNewMember.BorderStyle = psBorderSingle
    objNewMember.Expanded = False
    objNewMember.Key = Key
    objNewMember.Selected = False
    objNewMember.SelBorderStyle = psBorderSingle
    Index = mCol.Count + 1
    For Index = 1 To mCol.Count
        If LCase(Caption) < LCase(mCol(Index).Caption) Then
            Exit For
        End If
    Next
    objNewMember.Index = Index
    For i = Index To mCol.Count
        mCol(i).Index = i + 1
    Next
    ' the key property is the property caption
    If Index <= mCol.Count Then
        mCol.Add objNewMember, Key, Index
    Else
        mCol.Add objNewMember, Key
    End If
    ' add new property
    Parent.Parent.TriggerEvent "AddNewProperty", objNewMember, Parent
    ' add parent reference
    objNewMember.Init Parent
    ' return the object created
    Set Add = objNewMember
    ' clean it up
    Set objNewMember = Nothing
End Function

Public Property Get Item(Index As Variant) As TProperty
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(Index)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(Index As Variant)
    mCol.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Private Function GetDefaultValue(ValueType As psPropertyType) As Variant
    Select Case ValueType
    Case psInteger: GetDefaultValue = 0
    Case psLong: GetDefaultValue = 0
    Case psSingle: GetDefaultValue = 0
    Case psDouble: GetDefaultValue = 0
    Case psCurrency: GetDefaultValue = 0
    Case psDate: GetDefaultValue = Date
    Case psString: GetDefaultValue = ""
    Case psObject: GetDefaultValue = Nothing
    Case psBoolean: GetDefaultValue = False
    Case psDecimal: GetDefaultValue = 0
    Case psByte: GetDefaultValue = 0
    Case psFont: GetDefaultValue = Nothing
    Case psPicture: GetDefaultValue = Nothing
    Case psFile: GetDefaultValue = ""
    Case psColor: GetDefaultValue = 0
    Case psDropDownList: GetDefaultValue = 0
    Case psCombo: GetDefaultValue = 0
    Case psTime: GetDefaultValue = Time
    Case psLongText: GetDefaultValue = ""
    Case psFolder: GetDefaultValue = ""
    Case psDropDownCheckList: GetDefaultValue = ""
    End Select
End Function

Private Function GetValueType(Value As Variant) As psPropertyType
    Select Case VarType(Value)
        Case vbBoolean
            GetValueType = psBoolean
        Case vbLong
            GetValueType = psLong
        Case vbString
            GetValueType = psString
        Case vbDouble
            GetValueType = psDouble
        Case vbCurrency
            GetValueType = psCurrency
        Case vbString
            GetValueType = psString
        Case vbInteger
            GetValueType = psInteger
        Case vbSingle
            GetValueType = psSingle
        Case vbObject
            GetValueType = psObject
        Case Else
            GetValueType = psCustom
    End Select
End Function
'-- end code
