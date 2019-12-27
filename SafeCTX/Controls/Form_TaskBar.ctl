VERSION 5.00
Begin VB.UserControl Form_TaskBar 
   BackColor       =   &H00FF0000&
   CanGetFocus     =   0   'False
   ClientHeight    =   3525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   Enabled         =   0   'False
   ForwardFocus    =   -1  'True
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   1470
   ToolboxBitmap   =   "Form_TaskBar.ctx":0000
   Windowless      =   -1  'True
   Begin VB.Timer tmrDelayedInit 
      Left            =   120
      Top             =   360
   End
   Begin VB.Timer tmrCheckMouseOver 
      Left            =   120
      Top             =   1800
   End
   Begin VB.Timer tmrAppFocus 
      Left            =   120
      Top             =   1320
   End
   Begin VB.Timer tmrSliding 
      Left            =   120
      Top             =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "TaskBar"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "Form_TaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Problems:
'   "runs" while in IDE
'   can't be moved to the left, right, bottom, etc...
'   can't be positioned other than centered
'   when it moves down, it's kinda slow
'   the whole thing has too many hacks involving timers

' ########### API Calls #############
Private Type RECT
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
End Type
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, rectangle As RECT) As Long
'
Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long
'
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
        X As Long
        Y As Long
End Type
'
Private Declare Function GetForegroundWindow Lib "user32" () As Long


' ######### Events ###########
Event AppGotFocus()
Event AppLostFocus()
Event EndOpenUp()
Event EndCloseUp()
Event StartOpenUp()
Event StartCloseUp()
Event ChangeCloseUp()
Event ChangeOpenUp()
Event MouseOver()
Event MouseLeft()

' ########## Member Vars #######
Private mbActivated As Boolean

Private miScreenWidth As Integer
Private miScreenHeight As Integer
Private moForm As Form

Private mbSliderOut As Boolean
Private miSliderHowFar As Integer
Private miSliderChanging As Integer
Private mbHaveFocus As Boolean
Private mbMouseOver As Boolean

'Default Property Values:
Const m_def_NumSteps = 4
Const m_def_HangDown = 2
'Property Variables:
Dim m_NumSteps As Integer
Dim m_HangDown As Integer
'Event Declarations:




Private Sub UserControl_Initialize()
    tmrDelayedInit.Enabled = True
    tmrDelayedInit.Interval = 1
End Sub

Private Sub tmrDelayedInit_Timer()
    On Error GoTo NoForm
    Set moForm = UserControl.Parent
    On Error GoTo 0
    
    Call GetScreenResolution
    
    Call moForm.Move((miScreenWidth - moForm.Width) / 2, _
                m_HangDown * Screen.TwipsPerPixelY - moForm.Height)
                
    Call SetTopMost(moForm.hwnd)
    
    mbActivated = True
    
    tmrCheckMouseOver.Enabled = True
    tmrCheckMouseOver.Interval = 200
    
    tmrAppFocus.Enabled = True
    tmrAppFocus.Interval = 200
    
    tmrDelayedInit.Enabled = False
    
    Exit Sub
    
NoForm:
    MsgBox Err.Description, vbMsgBoxHelpButton, , Err.HelpFile, Err.HelpContext
    mbActivated = False
    
    tmrDelayedInit.Enabled = False
End Sub

Private Sub GetScreenResolution()
    Dim r As RECT
    Call GetWindowRect(GetDesktopWindow(), r)
    
    miScreenWidth = (r.x2 - r.x1) * Screen.TwipsPerPixelX
    miScreenHeight = (r.y2 - r.y1) * Screen.TwipsPerPixelY
End Sub

Private Sub SetTopMost(hwnd As Long)
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub

Private Sub SetSliderOut(bSO As Boolean)
    tmrSliding.Interval = 1
    
    If (bSO) Then           ' We're opening up
        If (mbSliderOut = False) Then
            RaiseEvent StartOpenUp
        ElseIf (miSliderChanging < 0) Then
            RaiseEvent ChangeOpenUp
        End If
            
        miSliderChanging = moForm.Height / m_NumSteps
        tmrSliding.Enabled = True
    Else                    ' We're closing up
        If (mbSliderOut = True) Then
            RaiseEvent StartCloseUp
        ElseIf (miSliderChanging > 0) Then
            RaiseEvent ChangeCloseUp
        End If
        
        miSliderChanging = -moForm.Height / m_NumSteps
        tmrSliding.Enabled = True
    End If
End Sub

Private Sub tmrSliding_Timer()
    Dim iToBeTop As Integer

    iToBeTop = moForm.Top + miSliderChanging

    If (iToBeTop >= 0) Then
        Call moForm.Move(moForm.Left, 0)
        mbSliderOut = True

        miSliderChanging = 0
        tmrSliding.Enabled = False

        RaiseEvent EndOpenUp

        Exit Sub
    ElseIf (iToBeTop <= m_HangDown * Screen.TwipsPerPixelY - moForm.Height) Then
        Call moForm.Move(moForm.Left, m_HangDown * Screen.TwipsPerPixelY - moForm.Height)
        mbSliderOut = False

        miSliderChanging = 0
        tmrSliding.Enabled = False

        RaiseEvent EndCloseUp

        Exit Sub
    End If

    Call moForm.Move(moForm.Left, iToBeTop)
End Sub

Private Sub tmrCheckMouseOver_Timer()
    Dim bThisMouseOver As Boolean
    
    Dim p As POINTAPI
    Call GetCursorPos(p)
    
    ' Check the screen coordinates of our window.  If it's not in ours, close ourselves up.
    Dim r As RECT
    Call GetWindowRect(moForm.hwnd, r)
    bThisMouseOver = (p.X >= r.x1 And p.X <= r.x2 And p.Y >= r.y1 And p.Y <= r.y2)
    If (bThisMouseOver <> mbMouseOver) Then
        mbMouseOver = bThisMouseOver
        
        If (mbMouseOver) Then           ' Just got the mouse over
            RaiseEvent MouseOver
            If (Not mbHaveFocus) Then _
                Call SetSliderOut(True)
        Else                            ' Just lost mouse over
            RaiseEvent MouseLeft
            If (Not mbHaveFocus) Then _
                Call SetSliderOut(False)
        End If
    End If
End Sub

Private Sub tmrAppFocus_Timer()
    Dim bThisHaveFocus As Boolean
    
    bThisHaveFocus = (GetForegroundWindow() = moForm.hwnd)
    
    ' We've just changed states
    If (bThisHaveFocus <> mbHaveFocus) Then
        mbHaveFocus = bThisHaveFocus
        
        If (mbHaveFocus) Then        ' Got focus
            RaiseEvent AppGotFocus
            Call SetSliderOut(True)
        Else                        ' Lost focus
            RaiseEvent AppLostFocus
            Call SetSliderOut(False)
        End If
    End If
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,4
Public Property Get NumSteps() As Integer
Attribute NumSteps.VB_Description = "The number of steps drawn while moving the taskbar down."
    NumSteps = m_NumSteps
End Property

Public Property Let NumSteps(ByVal New_NumSteps As Integer)
    m_NumSteps = New_NumSteps
    PropertyChanged "NumSteps"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,2
Public Property Get HangDown() As Integer
Attribute HangDown.VB_Description = "How many pixels will hang down into the screen."
    HangDown = m_HangDown
End Property

Public Property Let HangDown(ByVal New_HangDown As Integer)
    m_HangDown = New_HangDown
    PropertyChanged "HangDown"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_NumSteps = m_def_NumSteps
    m_HangDown = m_def_HangDown
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_NumSteps = PropBag.ReadProperty("NumSteps", m_def_NumSteps)
    m_HangDown = PropBag.ReadProperty("HangDown", m_def_HangDown)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("NumSteps", m_NumSteps, m_def_NumSteps)
    Call PropBag.WriteProperty("HangDown", m_HangDown, m_def_HangDown)
End Sub

Public Function IsTaskbarOut()
    If (mbSliderOut Or miSliderChanging <> 0) Then
        IsTaskbarOut = True
    Else
        IsTaskbarOut = False
    End If
End Function

Public Function IsTaskbarMoving()
    If (miSliderChanging <> 0) Then
        IsTaskbarMoving = True
    Else
        IsTaskbarMoving = False
    End If
End Function
