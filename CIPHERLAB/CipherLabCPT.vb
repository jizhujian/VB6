' ******************************************************************************************
'
' ******************************************************************************************
<Microsoft.VisualBasic.ComClass(CipherLabCPT.ClassId, CipherLabCPT.InterfaceId, CipherLabCPT.EventsId)> _
Public Class CipherLabCPT

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Const ClassId As String = "e74092e3-d42f-4d78-8221-45ad12735d6d"
  Public Const InterfaceId As String = "ffe6af03-b905-4032-bd62-22cf58094c82"
  Public Const EventsId As String = "7e5d53d1-7e49-463e-963b-a9b288525501"

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Private Declare Function CODEWARE_CPTOpenComPort Lib "CPTCommunication.dll" (ByVal portNumber As Integer, ByVal baudRate As Integer, ByVal portType As Integer) As Integer
  Private Declare Sub CODEWARE_CPTCloseComPort Lib "CPTCommunication.dll" ()
  Private Declare Function CODEWARE_CPTReadComPort Lib "CPTCommunication.dll" (ByVal data As String) As Integer
  Private Declare Sub CODEWARE_CPTWriteComPort Lib "CPTCommunication.dll" (ByVal data As String)
  Private Declare Sub CODEWARE_CPTWriteComPortBin Lib "CPTCommunication.dll" (ByVal data As Byte(), ByVal length As Integer)

  Private Declare Function CODEWARE_CPTStartDataUpload Lib "CPTCommunication.dll" (ByVal portNumber As Integer, ByVal baudRate As Integer, ByVal portType As Integer) As Integer
  Private Declare Function CODEWARE_CPTReadDataRecord Lib "CPTCommunication.dll" (ByVal data As String) As Integer
  Private Declare Sub CODEWARE_CPTFinishDataUpload Lib "CPTCommunication.dll" ()

  Private Declare Function CODEWARE_CPTStartLookupDownload Lib "CPTCommunication.dll" (ByVal portNumber As Integer, ByVal baudRate As Integer, ByVal portType As Integer) As Integer
  Private Declare Function CODEWARE_CPTSendLookupRecord Lib "CPTCommunication.dll" (ByVal data As String) As Integer
  Private Declare Sub CODEWARE_CPTFinishLookupDownload Lib "CPTCommunication.dll" ()

  Private Declare Function CODEWARE_CPTShowErrorMessage Lib "CPTCommunication.dll" (ByVal show As Integer) As Integer

  '*****************************************************************************************
  ' 端口类型
  '*****************************************************************************************
  Public Enum PortTypeEnum
    RS232 = 1
    CradleIr = 2
    IrDa = 3
  End Enum

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub New()
    MyBase.New()
    ShowErrorMessage(False)
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub OpenComPort(ByVal portName As String, Optional ByVal baudRate As Integer = 115200, _
    Optional ByVal portType As PortTypeEnum = PortTypeEnum.CradleIr)
    Dim returnValue As Integer = CODEWARE_CPTOpenComPort(CInt(portName.TrimStart("C"c, "O"c, "M"c)), baudRate, portType)
    If Not (returnValue = 0) Then
      HandleError(returnValue)
    End If
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub CloseComPort()
    CODEWARE_CPTCloseComPort()
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Function ReadComPort() As String
    Dim data As String = Microsoft.VisualBasic.StrDup(256, Microsoft.VisualBasic.vbNullChar)
    Dim returnValue As Integer = CODEWARE_CPTReadComPort(data)
    If (returnValue > 0) Then
      ReadComPort = data.Substring(0, returnValue)
    Else
      ReadComPort = ""
      If Not (returnValue = 0) Then
        HandleError(returnValue)
      End If
    End If
  End Function

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub WriteComPort(ByVal data As String)
    CODEWARE_CPTWriteComPort(data)
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub WriteComPortBin(ByVal data As Byte(), ByVal length As Integer)
    CODEWARE_CPTWriteComPortBin(data, length)
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub StartDataUpload(ByVal portName As String, Optional ByVal baudRate As Integer = 115200, _
    Optional ByVal portType As PortTypeEnum = PortTypeEnum.CradleIr)
    Dim returnValue As Integer = CODEWARE_CPTStartDataUpload(CInt(portName.TrimStart("C"c, "O"c, "M"c)), baudRate, portType)
    If Not (returnValue = 0) Then
      HandleError(returnValue)
    End If
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Function ReadDataRecord() As String
    Dim data As String = Microsoft.VisualBasic.StrDup(256, Microsoft.VisualBasic.vbNullChar)
    Dim returnValue As Integer = CODEWARE_CPTReadDataRecord(data)
    If (returnValue > 0) Then
      ReadDataRecord = data.Substring(0, returnValue)
    Else
      ReadDataRecord = ""
      If Not (returnValue = -5) Then
        HandleError(returnValue)
      End If
    End If
  End Function

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub FinishDataUpload()
    CODEWARE_CPTFinishDataUpload()
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub StartLookupDownload(ByVal portName As String, Optional ByVal baudRate As Integer = 115200, _
    Optional ByVal portType As PortTypeEnum = PortTypeEnum.CradleIr)
    Dim returnValue As Integer = CODEWARE_CPTStartLookupDownload(CInt(portName.TrimStart("C"c, "O"c, "M"c)), baudRate, portType)
    If Not (returnValue = 0) Then
      HandleError(returnValue)
    End If
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub SendLookupRecord(ByVal data As String)
    Dim returnValue As Integer = CODEWARE_CPTSendLookupRecord(data)
    If Not (returnValue = 0) Then
      HandleError(returnValue)
    End If
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub FinishLookupDownload()
    CODEWARE_CPTFinishLookupDownload()
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Sub ShowErrorMessage(ByVal show As Boolean)
    CODEWARE_CPTShowErrorMessage(CInt(Microsoft.VisualBasic.IIf(show, 1, 0)))
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Private Sub HandleError(ByVal returnValue As Integer)
    Select Case returnValue
      Case -1 ' ComPortNotOpen
        Throw New System.IO.IOException
      Case -2 ' CradleIrNotSet
        Throw New System.IO.IOException
      Case -3 ' ComPortNotInit
        Throw New System.IO.IOException
      Case -4 ' Timeout
        Throw New System.TimeoutException
      Case -5 ' NotMoreData
        Throw New System.IO.InvalidDataException
      Case -6 ' InvalidRecord
        Throw New System.ArgumentException
      Case Else
        Throw New System.Exception
    End Select
  End Sub

End Class
