VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStrings"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function RTF2Text(ByVal strRTF As String) As String
  Dim frm As New frmRTF
  frm.rtfTextBox.TextRTF = strRTF
  RTF2Text = frm.rtfTextBox.Text
  Unload frm
  Set frm = Nothing
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function Text2RTF(ByVal strText As String) As String
  Dim frm As New frmRTF
  frm.rtfTextBox.TextRTF = strText
  Text2RTF = frm.rtfTextBox.TextRTF
  Unload frm
  Set frm = Nothing
End Function

