VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IconResource"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function LoadResBitmap(ByVal ID As String) As IPictureDisp
  Set LoadResBitmap = LoadResPicture(ID, vbResBitmap)
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function LoadResIcon(ByVal ID As String) As IPictureDisp
Attribute LoadResIcon.VB_Description = "¼ÓÔØIcon×ÊÔ´"
  Set LoadResIcon = LoadResPicture(ID, vbResIcon)
End Function
