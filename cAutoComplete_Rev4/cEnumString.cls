VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEnumString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements oleexpimp.IEnumString
Private sItems() As String
Private nItems As Long
Private nCur As Long
Private iescnt As Long

Private mOldNext As Long
Private mOldSkip As Long

Public Function testcall()
Debug.Print "testcall success"
End Function
Private Sub Class_Initialize()

    Dim pVtable As oleexpimp.IEnumString
    Set pVtable = Me
    
    mOldNext = SwapVtableEntry(ObjPtr(pVtable), 4, AddressOf EnumStringNext)
    mOldSkip = SwapVtableEntry(ObjPtr(pVtable), 5, AddressOf EnumStringSkip)
    Debug.Print "IEnumString Swapped functions."
    
End Sub

Private Sub Class_Terminate()

    Dim pVtable As oleexpimp.IEnumString
    Set pVtable = Me
    
    mOldNext = SwapVtableEntry(ObjPtr(pVtable), 4, mOldNext)
    mOldSkip = SwapVtableEntry(ObjPtr(pVtable), 5, mOldSkip)
    
End Sub

Private Sub IEnumString_Next(ByVal celt As Long, ByVal rgelt As Long, ByVal pceltFetched As Long)
'Replaced

End Sub

Private Sub IEnumString_Skip(ByVal celt As Long)
'Replaced

End Sub
Private Sub IEnumString_Reset()
Debug.Print "IES_Reset"
StringCountReset
End Sub
Private Sub IEnumString_Clone(ppenum As IEnumString)
Debug.Print "IES_Clone"
Err.Raise E_NOTIMPL
End Sub
Public Function IES_Next(ByVal celt As Long, ByVal rgelt As Long, ByVal pceltFetched As Long) As Long
Debug.Print "cEnumString_Next"
Dim lpString As Long
Dim i As Long
Dim celtFetched As Long
If rgelt = 0 Then
    IES_Next = E_POINTER
    Exit Function
End If

For i = 0 To (celt - 1)
    If nCur = nItems Then Exit For
    lpString = CoTaskMemAlloc(LenB(sItems(nCur)) & vbNullChar)
    If lpString = 0 Then IES_Next = S_FALSE: Exit Function
    
    CopyMemory ByVal lpString, ByVal StrPtr(sItems(nCur)), LenB(sItems(nCur) & vbNullChar)
    CopyMemory ByVal UnsignedAdd(rgelt, i * 4), lpString, 4&
    
    nCur = nCur + 1
    celtFetched = celtFetched + 1
Next i
 If pceltFetched Then
    CopyMemory ByVal pceltFetched, celtFetched, 4&
 End If
 If i <> celt Then IES_Next = S_FALSE
 Debug.Print "IES_Next retval=" & IES_Next

End Function
Public Function IES_Skip(ByVal celt As Long) As Long
If nCur + celt <= nItems Then
    nCur = nCur + celt
    IES_Skip = S_OK
Else
    IES_Skip = S_FALSE
End If
End Function
Public Sub SetACStringList(sz() As String)
Dim i As Long
ReDim sItems(UBound(sz))
For i = 0 To UBound(sz)
    sItems(i) = sz(i)
Next i
nItems = UBound(sItems) + 1
End Sub
Public Sub ClearACStringList()
ReDim sItems(0)
nItems = 0
End Sub
Public Sub StringCountReset()
nCur = 0
End Sub
