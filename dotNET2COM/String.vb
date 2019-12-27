''' <summary>
''' �ַ���������
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass([String].ClassId, [String].InterfaceId, [String].EventsId)> _
Public Class [String]

  ''' <summary>
  ''' COMע�����
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "2745c3ad-d0d4-4784-bea6-9aa2e89ed5be"
  Public Const InterfaceId As String = "abe7946e-2666-454d-b9f5-63939a21c62f"
  Public Const EventsId As String = "1cf11a97-d0ab-40ce-a726-bfbc4434d46a"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' �� 8 λ�޷�������������ת��Ϊ���� Base64 ���ֱ���ĵ�Ч�ַ�����ʾ��ʽ
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BytesToBase64String(ByRef Value() As Byte) As String
    Return System.Convert.ToBase64String(Value)
  End Function

  ''' <summary>
  ''' ��ָ�����ַ������������������ݱ���Ϊ Base64 ���֣�ת��Ϊ��Ч�� 8 λ�޷�����������
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Base64StringToBytes(ByVal Value As String) As Byte()
    Return System.Convert.FromBase64String(Value)
  End Function

 
  ''' <summary>
  ''' ����ȫ��Ψһ��ʶ�� (GUID)
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function NewGuid() As String
    Return System.Guid.NewGuid.ToString
  End Function

  ''' <summary>
  ''' ��˫�ֽڣ�ȫ�ǣ��ַ���ת���ɵ��ֽڣ���ǣ��ַ���
  ''' </summary>
  ''' <param name="dbcString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function DBC2SBC(ByVal dbcString As String) As String
    Return Microsoft.VisualBasic.Strings.StrConv(dbcString, Microsoft.VisualBasic.VbStrConv.Narrow)
  End Function

  ''' <summary>
  ''' �����ֽڣ���ǣ��ַ���ת����˫�ֽڣ�ȫ�ǣ��ַ���
  ''' </summary>
  ''' <param name="sbcString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SBC2DBC(ByVal sbcString As String) As String
    Return Microsoft.VisualBasic.Strings.StrConv(sbcString, Microsoft.VisualBasic.VbStrConv.Wide)
  End Function

  ''' <summary>
  ''' ��ʽ����ʾ��ֵ
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="numDigitsAfterDecimal">С��λ��</param>
  ''' <param name="includeLeadingDigit">С����ǰ��ʾǰ����</param>
  ''' <param name="useParensForNegativeNumbers">����������</param>
  ''' <param name="groupDigits">��ʾǧ��λ�ָ���</param>
  ''' <param name="fixedDigitsAfterDecimal">С�������ʾ�̶�λ��</param>
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
  ''' �� URL �ַ������б���
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UrlEncode(ByVal value As String) As String
    Return System.Web.HttpUtility.UrlEncode(value)
  End Function

  ''' <summary>
  ''' ���Ѿ�Ϊ�� URL �д����������ַ���ת��Ϊ������ַ�����
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function UrlDecode(ByVal value As String) As String
    Return System.Web.HttpUtility.UrlDecode(value)
  End Function

  ''' <summary>
  ''' ���ַ���ת��Ϊ HTML ������ַ�����
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HtmlEncode(ByVal value As String) As String
    Return System.Web.HttpUtility.HtmlEncode(value)
  End Function

  ''' <summary>
  ''' ���Ѿ�Ϊ HTTP ������й� HTML ������ַ���ת��Ϊ�ѽ�����ַ�����
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HtmlDecode(ByVal value As String) As String
    Return System.Web.HttpUtility.HtmlDecode(value)
  End Function

End Class
