''' <summary>
''' 可变字符字符串
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(StringBuilder.ClassId, StringBuilder.InterfaceId, StringBuilder.EventsId)> _
Public Class StringBuilder

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "3bd02e47-eb3e-4140-b0c0-716a5255f221"
  Public Const InterfaceId As String = "cfc691ad-7e90-4a60-ae5a-c262ab19f09a"
  Public Const EventsId As String = "aa3ab712-e21c-49d3-8b8c-06d462e3d0e7"

  Dim sb As System.Text.StringBuilder

  Public Sub New()
    MyBase.New()
    sb = New System.Text.StringBuilder
  End Sub

  ''' <summary>
  ''' 在此实例的结尾追加指定字符串的副本。
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks></remarks>
  Public Sub Append(ByVal value As String)
    sb.Append(value)
  End Sub

  ''' <summary>
  ''' 将后面跟有默认行终止符的指定字符串的副本追加到当前 StringBuilder 对象的末尾。
  ''' </summary>
  ''' <param name="value"></param>
  ''' <remarks></remarks>
  Public Sub AppendLine(ByVal value As String)
    sb.AppendLine(value)
  End Sub

  ''' <summary>
  ''' 将此实例的值转换为 String
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Overrides Function ToString() As String
    Return sb.ToString
  End Function

  ''' <summary>
  ''' 从当前 StringBuilder 实例中移除所有字符。
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Clear()
    sb.Clear()
  End Sub

End Class
