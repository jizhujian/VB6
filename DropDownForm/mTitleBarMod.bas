Attribute VB_Name = "mTitleBarMod"
Option Explicit


Private Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_WNDPROC = (-4)

Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Const GW_OWNER = 4
Private Const GW_HWNDNEXT = 2

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_WINDOWPOSCHANGING = &H46
Private Const WM_NCACTIVATE = &H86
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_SHOWWINDOW = &H18

Private Const WS_VISIBLE = &H10000000


Public Sub AttachTitleBarMod(ByVal hwnd As Long)
Dim lhWndVBOwn As Long
Dim lSubClassWnd As Long
Dim lCount As Long

   lhWndVBOwn = GetWindow(hwnd, GW_OWNER)
   
   InstallWndProc lhWndVBOwn, hwnd, plAddressOf(AddressOf WndProc)
   
   lCount = GetProp(lhWndVBOwn, "AttachCount")
   lCount = lCount + 1
   SetProp lhWndVBOwn, "AttachCount", lCount
   SetProp lhWndVBOwn, "Attach" & lCount, hwnd
   SetProp lhWndVBOwn, "WndProc" & lCount, plAddressOf(AddressOf WndProc)
   
End Sub
Public Sub DetachTitleBarMod(ByVal hwnd As Long)
Dim lhWndVBOwn As Long
Dim lSubClassWnd As Long
Dim bNoSubclass As Boolean
Dim i As Long
Dim lIdx As Long
Dim lCount As Long

   lhWndVBOwn = GetWindow(hwnd, GW_OWNER)
   
   lSubClassWnd = GetProp(lhWndVBOwn, "SubclassWnd")
   If lSubClassWnd = hwnd Then
      SetProp lhWndVBOwn, "SubclassWnd", 0
      bNoSubclass = True
   End If
   
   lCount = GetProp(lhWndVBOwn, "AttachCount")
   For i = 1 To lCount
      If GetProp(lhWndVBOwn, "Attach" & i) = hwnd Then
         lIdx = i
         Exit For
      End If
   Next i
   
   If lCount = 1 Then
      ' Time to clear up
      RemoveProp lhWndVBOwn, "SubclassWnd"
      RemoveProp lhWndVBOwn, "AttachCount"
      RemoveProp lhWndVBOwn, "Attach1"
      RemoveProp lhWndVBOwn, "WndProc1"
      InstallWndProc lhWndVBOwn, 0, 0
   Else
      ' Still some left:
      For i = lIdx To lCount - 1
         SetProp lhWndVBOwn, "Attach" & i, GetProp(lhWndVBOwn, "Attach" & i + 1)
         SetProp lhWndVBOwn, "WndProc" & i, GetProp(lhWndVBOwn, "WndProc" & i + 1)
      Next i
      RemoveProp lhWndVBOwn, "Attach" & lCount
      RemoveProp lhWndVBOwn, "WndProc" & lCount
      lCount = lCount - 1
      SetProp lhWndVBOwn, "AttachCount", lCount
      
      If bNoSubclass Then
         ' Tx to hWnd1:
         InstallWndProc lhWndVBOwn, GetProp(lhWndVBOwn, "Attach1"), GetProp(lhWndVBOwn, "WndProc1")
      End If
      
   End If

End Sub
Private Sub InstallWndProc(ByVal hWndVBOwner As Long, ByVal hwnd As Long, ByVal lPtr As Long)
Dim lPtrOrig As Long
Dim iCount As Long, i As Long

   lPtrOrig = GetProp(hWndVBOwner, "OrigWndProc")
   
   If hwnd = 0 Then
      If Not (lPtrOrig = 0) Then
         ' Restore:
         Debug.Print "...Restoring Original WndProc"
         SetWindowLong hWndVBOwner, GWL_WNDPROC, lPtrOrig
      End If
      RemoveProp hWndVBOwner, "OrigWndProc"
      RemoveProp hWndVBOwner, "SubclassWnd"
      ' Normally we expect iCount to be zero here.
      ' However, this will ensure we can detach
      ' everything *regardless* of whether we are
      ' detaching in an order manner
      iCount = GetProp(hWndVBOwner, "AttachCount")
      Debug.Print "AttachCount:"; iCount
      For i = 1 To iCount
         RemoveProp hWndVBOwner, "Attach" & i
         RemoveProp hWndVBOwner, "WndProc" & i
      Next i
      Debug.Print "Cleared"
      
   Else
      Debug.Print "Setting WndProc"
      If lPtrOrig = 0 Then
         ' New subclass:
         Debug.Print "...Installing WndProc"
         lPtrOrig = SetWindowLong(hWndVBOwner, GWL_WNDPROC, lPtr)
         Debug.Print "...Storing Original WndProc", lPtrOrig
         SetProp hWndVBOwner, "OrigWndProc", lPtrOrig
         Debug.Print GetProp(hWndVBOwner, "OrigWndProc")
      End If
   End If
   
End Sub
Private Function plAddressOf(ByVal lPtr As Long) As Long
   plAddressOf = lPtr
End Function
Private Function WndProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lhWNdOwner As Long
Dim lhWnd As Long
Dim lPtr As Long
Dim lS As Long
Static bInHere As Boolean
   '
   Select Case iMsg
   Case WM_WINDOWPOSCHANGING
      WndProc = DefWindowProc(hwnd, iMsg, wParam, lParam)
      If Not bInHere Then
         bInHere = True
         lhWNdOwner = GetWindow(GetActiveWindow(), GW_OWNER)
         If lhWNdOwner = hwnd Then
            ' Top level:
            Debug.Print "TopLevel"
            If IsWindowVisible(hwnd) = 0 Then
               lS = GetWindowLong(hwnd, GWL_STYLE)
               SetWindowLong hwnd, GWL_STYLE, lS Or WS_VISIBLE
            End If
         Else
            ' Not top level:
            Debug.Print "NotTopLevel"
            ' Top level VB window:
            lhWnd = FindTopVBWindow(GetActiveWindow(), hwnd)
            If lhWnd <> 0 Then
               SendMessage lhWnd, WM_NCACTIVATE, 1, ByVal 0&
            End If
            lS = GetWindowLong(hwnd, GWL_STYLE)
            SetWindowLong hwnd, GWL_STYLE, lS And Not WS_VISIBLE
            
         End If
         bInHere = False
      End If
      
   Case Else
      WndProc = DefWindowProc(hwnd, iMsg, wParam, lParam)
   End Select
End Function
Private Function FindTopVBWindow(ByVal hWNdStart As Long, ByVal hWndVB As Long) As Long
Dim lhWnd As Long
   Do
      lhWnd = GetWindow(hWNdStart, GW_OWNER)
      If lhWnd = 0 Or lhWnd = hWndVB Then
         Exit Function
      Else
         FindTopVBWindow = lhWnd
         hWNdStart = lhWnd
      End If
   Loop
End Function


