Attribute VB_Name = "mACHelper"
Option Explicit
Public IES_RetVal As Long
Private sItems() As String
Private nItems As Long
Private nCur As Long
Private iescnt As Long
Public Declare Function vbaObjSetAddRef Lib "msvbvm60.dll" Alias "__vbaObjSetAddref" (ByRef objDest As Object, ByVal pObject As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function VirtualProtect Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Const PAGE_EXECUTE_READWRITE As Long = &H40&

Public Const E_NOTIMPL = &H80004001      '_HRESULT_TYPEDEF_(0x80004001L)
Public Const E_OUTOFMEMORY = &H8007000E  '_HRESULT_TYPEDEF_(0x8007000EL)
Public Const E_INVALIDARG = &H80070057   '_HRESULT_TYPEDEF_(0x80070057L)
Public Const E_NOINTERFACE = &H80004002  '_HRESULT_TYPEDEF_(0x80004002L)
Public Const E_POINTER = &H80004003      '_HRESULT_TYPEDEF_(0x80004003L)
Public Const E_HANDLE = &H80070006       '_HRESULT_TYPEDEF_(0x80070006L)
Public Const E_ABORT = &H80004004        '_HRESULT_TYPEDEF_(0x80004004L)
Public Const E_FAIL = &H80004005         '_HRESULT_TYPEDEF_(0x80004005L)
Public Const E_ACCESSDENIED = &H80070005 '_HRESULT_TYPEDEF_(0x80070005L)
Public Function SwapVtableEntry(pObj As Long, EntryNumber As Integer, ByVal lpfn As Long) As Long

    Dim lOldAddr As Long
    Dim lpVtableHead As Long
    Dim lpfnAddr As Long
    Dim lOldProtect As Long

    CopyMemory lpVtableHead, ByVal pObj, 4
    lpfnAddr = lpVtableHead + (EntryNumber - 1) * 4
    CopyMemory lOldAddr, ByVal lpfnAddr, 4

    Call VirtualProtect(lpfnAddr, 4, PAGE_EXECUTE_READWRITE, lOldProtect)
    CopyMemory ByVal lpfnAddr, lpfn, 4
    Call VirtualProtect(lpfnAddr, 4, lOldProtect, lOldProtect)

    SwapVtableEntry = lOldAddr

End Function
Public Function UnsignedAdd(ByVal Start As Long, ByVal Incr As Long) As Long
UnsignedAdd = ((Start Xor &H80000000) + Incr) Xor &H80000000
End Function

Public Function EnumStringNext(ByVal this As Long, ByVal celt As Long, ByVal rgelt As Long, ByVal pceltFetched As Long) As Long
'Debug.Print "esn ptr=" & this
Dim cObj As cEnumString
vbaObjSetAddRef cObj, this
If (cObj Is Nothing) = False Then
    Debug.Print "obj set"
'    cObj.testcall
    EnumStringNext = cObj.IES_Next(celt, rgelt, pceltFetched)
Else
    Debug.Print "esn obj fail"
    EnumStringNext = S_FALSE
End If

End Function
Public Function EnumStringSkip(ByVal this As Long, ByVal celt As Long) As Long
'Debug.Print "ess ptr=" & this
Dim cObj As cEnumString
vbaObjSetAddRef cObj, this
If (cObj Is Nothing) = False Then
    Debug.Print "obj set"
    EnumStringSkip = cObj.IES_Skip(celt)
Else
    Debug.Print "ess obj fail"
End If

End Function


