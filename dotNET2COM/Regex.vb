''' <summary>
''' 正则表达式函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Regex.ClassId, Regex.InterfaceId, Regex.EventsId)> _
Public Class Regex

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "6bcd5e7d-0128-4144-809e-7d7f4ec0e988"
  Public Const InterfaceId As String = "76eb7189-3424-4f68-9fcf-dd2160144a5d"
  Public Const EventsId As String = "152470d8-00ee-4209-8180-b1e6d2a2342e"

  Public Sub New()
    MyBase.New()
  End Sub

  Private _pattern As String = ""
  Private _ignoreCase As Boolean = True

  ''' <summary>
  ''' 要匹配的正则表达式模式
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Pattern() As String
    Get
      Return _pattern
    End Get
    Set(ByVal value As String)
      _pattern = value
    End Set
  End Property

  ''' <summary>
  ''' 指定不区分大小写的匹配
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property IgnoreCase() As Boolean
    Get
      Return _ignoreCase
    End Get
    Set(ByVal value As Boolean)
      _ignoreCase = value
    End Set
  End Property

  ''' <summary>
  ''' 使用指定的匹配选项在指定的输入字符串中搜索指定的正则表达式的所有匹配项
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Matches(ByVal value As String) As String()
    Dim options As System.Text.RegularExpressions.RegexOptions
    If IgnoreCase Then
      options = System.Text.RegularExpressions.RegexOptions.IgnoreCase
    Else
      options = System.Text.RegularExpressions.RegexOptions.None
    End If
    Dim matchList As System.Text.RegularExpressions.MatchCollection
    matchList = System.Text.RegularExpressions.Regex.Matches(value, Pattern, options)
    Dim matchValues(matchList.Count - 1) As String
    For i As Integer = 0 To matchList.Count - 1
      matchValues(i) = matchList(i).Value
    Next
    Return matchValues
  End Function

  ''' <summary>
  ''' 使用指定的匹配选项在指定的输入字符串中搜索指定的正则表达式的所有匹配项，返回正则表达式中的命名值对
  ''' 正则表达式必须包含&lt;name[0-99]&gt;&lt;value[0-99]&gt;
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ParseNameValues(ByVal value As String) As String()
    Dim nameList As System.Text.RegularExpressions.MatchCollection
    nameList = System.Text.RegularExpressions.Regex.Matches(Pattern, "\(\?\<(?<name>name\d*)\>", _
      System.Text.RegularExpressions.RegexOptions.IgnoreCase)
    Dim valueList As System.Text.RegularExpressions.MatchCollection
    valueList = System.Text.RegularExpressions.Regex.Matches(Pattern, "\(\?\<(?<value>value\d*)\>", _
      System.Text.RegularExpressions.RegexOptions.IgnoreCase)
    Dim options As System.Text.RegularExpressions.RegexOptions
    If IgnoreCase Then
      options = System.Text.RegularExpressions.RegexOptions.IgnoreCase
    Else
      options = System.Text.RegularExpressions.RegexOptions.None
    End If
    Dim matchList As System.Text.RegularExpressions.MatchCollection
    matchList = System.Text.RegularExpressions.Regex.Matches(value, Pattern, options)
    Dim matchValues(matchList.Count * (nameList.Count + valueList.Count) - 1) As String
    For i As Integer = 0 To matchList.Count - 1
      With matchList(i).Groups
        'matchValues(i * (nameList.Count + valueList.Count)) = ""
        For j As Integer = 0 To nameList.Count - 1
          matchValues(i * (nameList.Count + valueList.Count) + j) = .Item(nameList(j).Groups("name").Value).Value
        Next
        For j As Integer = 0 To valueList.Count - 1
          matchValues(i * (nameList.Count + valueList.Count) + nameList.Count + j) = .Item(valueList(j).Groups("value").Value).Value
        Next
      End With
    Next
    Return matchValues
  End Function

  ''' <summary>
  ''' 在指定的输入字符串内，使用指定的替换字符串替换与指定正则表达式匹配的所有字符串。指定的选项将修改匹配操作
  ''' </summary>
  ''' <remarks></remarks>
  Public Function Replace(ByVal value As String, ByVal replacement As String) As String
    Dim options As System.Text.RegularExpressions.RegexOptions
    If IgnoreCase Then
      options = System.Text.RegularExpressions.RegexOptions.IgnoreCase
    Else
      options = System.Text.RegularExpressions.RegexOptions.None
    End If
    Return System.Text.RegularExpressions.Regex.Replace(value, _pattern, replacement, options)
  End Function

End Class
