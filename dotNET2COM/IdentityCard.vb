''' <summary>
''' 居民身份证号码函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(IdentityCard.ClassId, IdentityCard.InterfaceId, IdentityCard.EventsId)> _
Public Class IdentityCard

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "1df9b730-a394-4d60-a729-9c117dfb8f99"
  Public Const InterfaceId As String = "08f2c000-f242-4229-bb96-7b36274d1ec1"
  Public Const EventsId As String = "dc34b504-aaa4-47af-a624-1392212980c2"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 性别类型
  ''' </summary>
  ''' <remarks></remarks>
  Public Enum GenderEnum
    ''' <summary>
    ''' 女
    ''' </summary>
    ''' <remarks></remarks>
    Women
    ''' <summary>
    ''' 男
    ''' </summary>
    ''' <remarks></remarks>
    Men
  End Enum

  Private _coding As String
  Private _birthday As Date
  Private _gender As GenderEnum


  ''' <summary>
  ''' 居民身份证号码
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Coding() As String
    Get
      Return _coding
    End Get
    Set(ByVal value As String)
      Dim tempCoding As String
      Dim tempBirthday As Date
      Dim tempGender As GenderEnum
      tempCoding = value.ToUpper
      If (tempCoding.Length = 15) Then
        tempCoding = tempCoding.Insert(6, "19")
        tempCoding &= CalculateParityBit(tempCoding)
      ElseIf (tempCoding.Length = 18) Then
        If (tempCoding.Chars(17) <> CalculateParityBit(tempCoding)) Then
          Throw New System.ArgumentException("居民身份证号码校验码错误。")
        End If
      Else
        Throw New System.ArgumentException("居民身份证号码必须是15位或18位。")
      End If
      Try
        tempBirthday = New Date(CInt(tempCoding.Substring(6, 4)), CInt(tempCoding.Substring(10, 2)), CInt(tempCoding.Substring(12, 2)))
      Catch
        Throw New System.ArgumentException("居民身份证号码中的生日错误。")
      End Try
      Try
        tempGender = CType(CInt(tempCoding.Substring(16, 1)) Mod 2, GenderEnum)
      Catch
        Throw New System.ArgumentException("居民身份证号码中的性别错误。")
      End Try
      _coding = tempCoding
      _birthday = tempBirthday
      _gender = tempGender
    End Set
  End Property

  ''' <summary>
  ''' 出生日期
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property Birthday() As Date
    Get
      Return _birthday
    End Get
  End Property

  ''' <summary>
  ''' 性别
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property Gender() As GenderEnum
    Get
      Return _gender
    End Get
  End Property

  ''' <summary>
  ''' 计算居民身份证最后一位校验码
  ''' </summary>
  ''' <param name="coding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CalculateParityBit(ByVal coding As String) As String
    Try
      Dim powerFactors As Integer() = New Integer() {7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2, 1}
      Dim power As Integer
      For i As Integer = 0 To 16
        power += CInt(coding.Substring(i, 1)) * powerFactors(i)
      Next
      Return "10X98765432".Chars(power Mod 11)
    Catch
      Throw New System.ArgumentException("身份证号码错误。")
    End Try
  End Function

End Class
