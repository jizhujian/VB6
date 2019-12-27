VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event OnTimer()

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private mblnEnabled As Boolean
Private mlngInterval As Long
Private mlngIntervalA As Long
Private mlngTickCount As Long

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get Enabled() As Boolean
  Enabled = mblnEnabled
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Let Enabled(blnEnabled As Boolean)
  mblnEnabled = blnEnabled
  If blnEnabled Then
    mlngTickCount = GetTickCount
    TimerLoop
  End If
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get Interval() As Long
  Interval = mlngInterval
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Let Interval(lngInterval As Long)
  mlngInterval = lngInterval
  mlngIntervalA = lngInterval * 1000
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Sub TimerLoop()
  Do While Not mblnEnabled = False
    If GetTickCount - mlngTickCount >= mlngIntervalA Then
      RaiseEvent OnTimer
      mlngTickCount = GetTickCount
    ElseIf GetTickCount = 0 Then
      mlngTickCount = 0
    ElseIf GetTickCount < mlngTickCount Then
      mlngTickCount = 0
    End If
    DoEvents
  Loop
End Sub
