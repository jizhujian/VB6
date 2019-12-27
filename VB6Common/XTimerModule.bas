Attribute VB_Name = "XTimerModule"
Option Explicit

Public XTimerColl As New VBA.Collection

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Sub XTimeProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
  Dim objPtr2Object As Ptr2Object
  Dim objXTimer As XTimer
  Dim lpTimer As Long
  lpTimer = XTimerColl("ID:" & idEvent)
  Set objPtr2Object = New Ptr2Object
  Set objXTimer = objPtr2Object.ObjectFromPtr(lpTimer)
  Set objPtr2Object = Nothing
  objXTimer.PulseTimer
  Set objXTimer = Nothing
End Sub

