''' <summary>
''' ������ʽ������
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Regex.ClassId, Regex.InterfaceId, Regex.EventsId)> _
Public Class Regex

  ''' <summary>
  ''' COMע�����
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
  ''' Ҫƥ���������ʽģʽ
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
  ''' ָ�������ִ�Сд��ƥ��
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
  ''' ʹ��ָ����ƥ��ѡ����ָ���������ַ���������ָ����������ʽ������ƥ����
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
  ''' ʹ��ָ����ƥ��ѡ����ָ���������ַ���������ָ����������ʽ������ƥ�������������ʽ�е�����ֵ��
  ''' ������ʽ�������&lt;name[0-99]&gt;&lt;value[0-99]&gt;
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
  ''' ��ָ���������ַ����ڣ�ʹ��ָ�����滻�ַ����滻��ָ��������ʽƥ��������ַ�����ָ����ѡ��޸�ƥ�����
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
