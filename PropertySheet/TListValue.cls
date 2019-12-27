VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TListValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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
' 类模块       : TListValue.cls
' 作      者   : Marclei V Silva (MVS)
' 程序员       : Marclei V Silva (MVS) [Spnorte Consultoria de Inform醫ica]
' 编 写 日 期  : 06/16/2000 -- 08:33:11
' 输  入       : N/A
' 输  出       : N/A
' 描   述      : Simple list value information
'              :
'              :
' *******************************************************
Option Explicit

Private m_Caption As String
Private m_Value As Variant
Private m_ParentPtr As Long

Public Sub Init(ByVal Object As TProperty)
    m_ParentPtr = PtrFromObject(Object)
End Sub

Public Property Get Parent() As TProperty
    Set Parent = ObjectFromPtr(m_ParentPtr)
End Property

Public Property Let Value(ByVal vData As Variant)
    m_Value = vData
    If m_ParentPtr <> 0 Then
        Parent.Parent.Parent.TriggerEvent "ValueChanged", Parent, Parent.Row, vData
    End If
End Property

Public Property Set Value(ByVal vData As Variant)
    Set m_Value = vData
    If m_ParentPtr <> 0 Then
        Parent.Parent.Parent.TriggerEvent "ValueChanged", Parent, Parent.Row, vData
    End If
End Property

Public Property Get Value() As Variant
    If IsObject(m_Value) Then
        On Error Resume Next
        Set Value = m_Value
    Else
        Value = m_Value
    End If
End Property

Public Property Let Caption(ByVal vData As String)
    m_Caption = vData
    If m_ParentPtr <> 0 Then
        Parent.Parent.Parent.TriggerEvent "ValueChanged", Parent, Parent.Row, vData
    End If
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property
