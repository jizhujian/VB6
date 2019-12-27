VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cClipboard"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal Format As Long, ByVal hMem As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal flags As Long, ByVal length As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSource As Long, ByVal length As Long)

Private Const CF_TEXT = 1
Private Const CF_UNICODETEXT = 13
Private Const GHND = &H42&

Public Sub SetText(ByVal hwnd As Long, ByVal str As String)

  Dim s As String
  Dim hMem As Long
  Dim pMem As Long
  Dim bOk As Boolean

  If OpenClipboard(hwnd) = 0 Then Exit Sub

  EmptyClipboard

  If (str > "") Then

    ' ANSI
    s = StrConv(str, vbFromUnicode)
    hMem = GlobalAlloc(GHND, LenB(s) + 1)
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, ByVal StrPtr(s), LenB(s)
    GlobalUnlock hMem
    bOk = SetClipboardData(CF_TEXT, hMem) <> 0
    If Not bOk Then GlobalFree hMem

    ' UNICODE:
    hMem = GlobalAlloc(GHND, LenB(str) + 2)
    pMem = GlobalLock(hMem)
    CopyMemory ByVal pMem, ByVal StrPtr(str), LenB(str)
    GlobalUnlock hMem
    bOk = SetClipboardData(CF_UNICODETEXT, hMem) <> 0
    If Not bOk Then GlobalFree hMem

  End If

  CloseClipboard

End Sub

