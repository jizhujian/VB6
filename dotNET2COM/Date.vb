''' <summary>
''' 日期函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass([Date].ClassId, [Date].InterfaceId, [Date].EventsId)> _
Public Class [Date]

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "f7da544b-cdb1-47a1-a048-984b47a69a8c"
  Public Const InterfaceId As String = "3b1375f6-a3ae-4c79-b9a6-4667c995f4df"
  Public Const EventsId As String = "d907f05e-18bb-4bc0-867a-7c11cc8a9799"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 是否表示一个有效的日期值
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function IsDate(ByVal value As String) As Boolean
    Return Microsoft.VisualBasic.Information.IsDate(value)
  End Function

  ''' <summary>
  ''' 将字符串转换成日期值
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CDateA(ByVal value As String) As Date
    Return CDate(value)
  End Function

End Class
