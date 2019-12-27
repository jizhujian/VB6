''' <summary>
''' 字符串函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass([String].ClassId, [String].InterfaceId, [String].EventsId)> _
Public Class [String]

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "2745c3ad-d0d4-4784-bea6-9aa2e89ed5be"
  Public Const InterfaceId As String = "abe7946e-2666-454d-b9f5-63939a21c62f"
  Public Const EventsId As String = "1cf11a97-d0ab-40ce-a726-bfbc4434d46a"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 将 8 位无符号整数的数组转换为其用 Base64 数字编码的等效字符串表示形式
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BytesToBase64String(ByRef Value() As Byte) As String
    Return System.Convert.ToBase64String(Value)
  End Function

  ''' <summary>
  ''' 将指定的字符串（它将二进制数据编码为 Base64 数字）转换为等效的 8 位无符号整数数组
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Base64StringToBytes(ByVal Value As String) As Byte()
    Return System.Convert.FromBase64String(Value)
  End Function

 
  ''' <summary>
  ''' 返回全局唯一标识符 (GUID)
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function NewGuid() As String
    Return System.Guid.NewGuid.ToString
  End Function

  ''' <summary>
  ''' 将双字节（全角）字符串转换成单字节（半角）字符串
  ''' </summary>
  ''' <param name="dbcString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DBC2SBC(ByVal dbcString As String) As String
    Return Microsoft.VisualBasic.Strings.StrConv(dbcString, Microsoft.VisualBasic.VbStrConv.Narrow)
  End Function

  ''' <summary>
  ''' 将单字节（半角）字符串转换成双字节（全角）字符串
  ''' </summary>
  ''' <param name="sbcString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SBC2DBC(ByVal sbcString As String) As String
    Return Microsoft.VisualBasic.Strings.StrConv(sbcString, Microsoft.VisualBasic.VbStrConv.Wide)
  End Function

  ''' <summary>
  ''' 格式化显示数值
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="numDigitsAfterDecimal">小数位数</param>
  ''' <param name="includeLeadingDigit">小数点前显示前导零</param>
  ''' <param name="useParensForNegativeNumbers">负数用括号</param>
  ''' <param name="groupDigits">显示千分位分隔符</param>
  ''' <param name="fixedDigitsAfterDecimal">小数点后显示固定位数</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FormatNumber(ByVal value As Double, _
    Optional ByVal numDigitsAfterDecimal As Integer = 6, _
    Optional ByVal includeLeadingDigit As Boolean = True, _
    Optional ByVal useParensForNegativeNumbers As Boolean = False, _
    Optional ByVal groupDigits As Boolean = False, _
    Optional ByVal fixedDigitsAfterDecimal As Boolean = False) As String
    Dim formatString As String
    If numDigitsAfterDecimal > 0 Then
      formatString = "." & Microsoft.VisualBasic.StrDup(numDigitsAfterDecimal, CStr(Microsoft.VisualBasic.IIf(fixedDigitsAfterDecimal, "0", "#")))
    Else
      formatString = ""
    End If
    If includeLeadingDigit Then
      formatString = "0" & formatString
    Else
      formatString = "#" & formatString
    End If
    If groupDigits Then
      formatString = "#,##" & formatString
    End If
    formatString = formatString & ";" & CStr(Microsoft.VisualBasic.IIf(useParensForNegativeNumbers, "(", "-")) & _
      formatString & CStr(Microsoft.VisualBasic.IIf(useParensForNegativeNumbers, ")", "")) & ";#"
    Return value.ToString(formatString)
  End Function

  ''' <summary>
  ''' 对 URL 字符串进行编码
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UrlEncode(ByVal value As String) As String
    Return System.Web.HttpUtility.UrlEncode(value)
  End Function

  ''' <summary>
  ''' 将已经为在 URL 中传输而编码的字符串转换为解码的字符串。
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UrlDecode(ByVal value As String) As String
    Return System.Web.HttpUtility.UrlDecode(value)
  End Function

  ''' <summary>
  ''' 将字符串转换为 HTML 编码的字符串。
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HtmlEncode(ByVal value As String) As String
    Return System.Web.HttpUtility.HtmlEncode(value)
  End Function

  ''' <summary>
  ''' 将已经为 HTTP 传输进行过 HTML 编码的字符串转换为已解码的字符串。
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HtmlDecode(ByVal value As String) As String
    Return System.Web.HttpUtility.HtmlDecode(value)
  End Function

End Class
