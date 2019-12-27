VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function GenErrSource(ByVal strModuleName As String, ByVal strProcedureName As String) As String
  GenErrSource = Err.Source & _
    IIf(StrComp(Left$(strModuleName, (InStr(1, strModuleName, ".") - 1)), Err.Source, vbTextCompare) = 0, _
    Mid$(strModuleName, InStr(1, strModuleName, ".")), _
    " -> " & strModuleName) & _
    "." & strProcedureName
End Function


