''' <summary>
''' 串行端口函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(SerialPort.ClassId, SerialPort.InterfaceId, SerialPort.EventsId)> _
Public Class SerialPort

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "bc321f21-1b00-46dd-8440-2416ee48240f"
  Public Const InterfaceId As String = "8ac69c9a-3cbc-447d-8f35-96bf472966d3"
  Public Const EventsId As String = "51c7207e-02bd-4af5-ad3c-257118fffab6"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 获取当前计算机的串行端口名称数组
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSerialPortNames() As String()
    Return System.IO.Ports.SerialPort.GetPortNames
  End Function

End Class
