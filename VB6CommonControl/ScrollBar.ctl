VERSION 5.00
Begin VB.UserControl ScrollBar 
   Alignable       =   -1  'True
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8010
   ControlContainer=   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   8010
   ToolboxBitmap   =   "ScrollBar.ctx":0000
End
Attribute VB_Name = "ScrollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents msb As cScrollBars
Attribute msb.VB_VarHelpID = -1
Private mlngChildObjPtr As Long

Public Property Get Child() As Object
  Set Child = ObjectFromPtr(mlngChildObjPtr)
End Property

Public Property Set Child(ByRef Obj As Object)
  mlngChildObjPtr = ObjPtr(Obj)
  Obj.Move 0, 0
  UserControl_Resize
  PropertyChanged "Child"
End Property

Public Sub RecalcLayout()
  UserControl_Resize
End Sub

Private Sub UserControl_Initialize()
  Set msb = New cScrollBars
  msb.Create hWnd
End Sub

Private Sub UserControl_Terminate()
  Set msb = Nothing
End Sub

Private Sub UserControl_Resize()

  Dim ctl As Object
  Dim lHeight As Long
  Dim lWidth As Long
  Dim lProportion As Long

  If mlngChildObjPtr = 0 Then
    Exit Sub
  End If
  Set ctl = Child

  ' Pixels are the minimum change size for a screen object.
  ' Therefore we set the scroll bars in pixels.

  lHeight = (ctl.Height - ScaleHeight) \ Screen.TwipsPerPixelY
  If (lHeight > 0) Then
    lProportion = lHeight \ (ScaleHeight \ Screen.TwipsPerPixelY) + 1
    msb.LargeChange(efsVertical) = lHeight \ lProportion
    msb.Max(efsVertical) = lHeight
    msb.Visible(efsVertical) = True
    If (ctl.Height + ctl.Top) < ScaleHeight Then
      ctl.Top = ScaleHeight - ctl.Height
    End If
  Else
    msb.Visible(efsVertical) = False
    ctl.Top = 0
  End If

  lWidth = (ctl.Width - ScaleWidth) \ Screen.TwipsPerPixelX
  If (lWidth > 0) Then
    lProportion = lWidth \ (ScaleWidth \ Screen.TwipsPerPixelX) + 1
    msb.LargeChange(efsHorizontal) = lWidth \ lProportion
    msb.Max(efsHorizontal) = lWidth
    msb.Visible(efsHorizontal) = True
    If (ctl.Width + ctl.Left) < ScaleWidth Then
      ctl.Left = ScaleWidth - ctl.Width
    End If
  Else
    msb.Visible(efsHorizontal) = False
    ctl.Left = 0
  End If

End Sub

Private Sub msb_Change(eBar As EFSScrollBarConstants)
  Dim ctl As Control
  Set ctl = Child
  If (msb.Visible(eBar)) Then
    If (eBar = efsHorizontal) Then
      ctl.Left = -msb.Value(eBar) * Screen.TwipsPerPixelX
    Else
      ctl.Top = -msb.Value(eBar) * Screen.TwipsPerPixelY
    End If
  Else
    ctl.Move 0, 0
  End If
End Sub

Private Sub msb_Scroll(eBar As EFSScrollBarConstants)
  msb_Change eBar
End Sub
