VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cArray"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function ArrayIsInitialized(ByVal arr As Variant) As Boolean
  Dim i As Long
  On Error GoTo HERROR
  If IsArray(arr) Then
    i = LBound(arr)
    ArrayIsInitialized = True
  End If
  Exit Function
HERROR:
End Function
