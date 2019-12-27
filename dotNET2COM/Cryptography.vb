''' <summary>
''' 加密函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Cryptography.ClassId, Cryptography.InterfaceId, Cryptography.EventsId)> _
Public Class Cryptography

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "9e3c3a84-22cf-4bcd-b285-a712cb1eb94f"
  Public Const InterfaceId As String = "15285176-7117-4a5f-8942-0500f4f9a550"
  Public Const EventsId As String = "f4fdeb34-33b4-46d6-940d-5f59513116f9"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 计算字节数组的MD5哈希值
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ComputeByteArrayHash(ByRef Value() As Byte) As Byte()
    Return (New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(Value)
  End Function

  ''' <summary>
  ''' 计算字符串的MD5哈希值
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ComputeStringHash(ByVal Value As String, Optional ByVal textCoding As File.TextEncoding = File.TextEncoding.Default) As Byte()
    Return ComputeByteArrayHash(File.GetTextEncoding(textCoding).GetBytes(Value))
  End Function

  ''' <summary>
  ''' 加密字符串
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <param name="Key"></param>
  ''' <param name="IV"></param>
  ''' <param name="ValueTextEncoding"></param>
  ''' <param name="KeyTextEncoding"></param>
  ''' <param name="IVTextEncoding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Encrypt(ByVal Value As String, ByVal Key As String, Optional ByVal IV As String = "",
    Optional ByVal ValueTextEncoding As File.TextEncoding = File.TextEncoding.Default,
    Optional ByVal KeyTextEncoding As File.TextEncoding = File.TextEncoding.Default,
    Optional ByVal IVTextEncoding As File.TextEncoding = File.TextEncoding.Default) As String

    Dim rgbKey As Byte() = ComputeStringHash(Key, KeyTextEncoding)
    Dim rgbIV As Byte()
    If IV > "" Then
      rgbIV = ComputeStringHash(IV, IVTextEncoding)
    Else
      rgbIV = ComputeStringHash(Key, IVTextEncoding)
      System.Array.Reverse(rgbIV)
    End If
    Dim memoryStream As New System.IO.MemoryStream
    Dim rmCrypto As New System.Security.Cryptography.RijndaelManaged
    Dim cryptStream As New System.Security.Cryptography.CryptoStream(memoryStream, _
      rmCrypto.CreateEncryptor(rgbKey, rgbIV), System.Security.Cryptography.CryptoStreamMode.Write)
    Dim buffers As Byte() = File.GetTextEncoding(ValueTextEncoding).GetBytes(Value)
    cryptStream.Write(buffers, 0, buffers.Length)
    cryptStream.FlushFinalBlock()
    Dim EncryptBytes As Byte() = memoryStream.ToArray
    cryptStream.Close()
    memoryStream.Close()
    Return System.Convert.ToBase64String(EncryptBytes)

  End Function

  ''' <summary>
  ''' 解密字符串
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <param name="Key"></param>
  ''' <param name="IV"></param>
  ''' <param name="ValueTextEncoding"></param>
  ''' <param name="KeyTextEncoding"></param>
  ''' <param name="IVTextEncoding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Decrypt(ByRef Value As String, ByVal Key As String, Optional ByVal IV As String = "",
    Optional ByVal ValueTextEncoding As File.TextEncoding = File.TextEncoding.Default,
    Optional ByVal KeyTextEncoding As File.TextEncoding = File.TextEncoding.Default,
    Optional ByVal IVTextEncoding As File.TextEncoding = File.TextEncoding.Default) As String

    Dim rgbKey As Byte() = ComputeStringHash(Key, KeyTextEncoding)
    Dim rgbIV As Byte()
    If IV > "" Then
      rgbIV = ComputeStringHash(IV, IVTextEncoding)
    Else
      rgbIV = ComputeStringHash(Key, IVTextEncoding)
      System.Array.Reverse(rgbIV)
    End If
    Dim memoryStream As New System.IO.MemoryStream
    Dim rmCrypto As New System.Security.Cryptography.RijndaelManaged
    Dim cryptStream As New System.Security.Cryptography.CryptoStream(memoryStream, _
      rmCrypto.CreateDecryptor(rgbKey, rgbIV), System.Security.Cryptography.CryptoStreamMode.Write)
    Dim buffers As Byte() = System.Convert.FromBase64String(Value)
    cryptStream.Write(buffers, 0, buffers.Length)
    cryptStream.FlushFinalBlock()
    Dim decryptString As String = File.GetTextEncoding(ValueTextEncoding).GetString(memoryStream.ToArray)
    cryptStream.Close()
    memoryStream.Close()
    Return decryptString

  End Function

  ''' <summary>
  ''' 创建私钥
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePrivateKey() As String
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    Return System.Convert.ToBase64String(rsa.ExportCspBlob(True))
  End Function

  ''' <summary>
  ''' 提取公钥
  ''' </summary>
  ''' <param name="privateKey"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPublicKey(ByVal privateKey As String) As String
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    rsa.ImportCspBlob(System.Convert.FromBase64String(privateKey))
    Return System.Convert.ToBase64String(rsa.ExportCspBlob(False))
  End Function

  ''' <summary>
  ''' 创建数字签名
  ''' </summary>
  ''' <param name="privateKey">私钥</param>
  ''' <param name="value">需要签名的内容</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateSignature(ByVal privateKey As String, ByVal value As String) As String
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    rsa.ImportCspBlob(System.Convert.FromBase64String(privateKey))
    Dim formatter As New System.Security.Cryptography.RSAPKCS1SignatureFormatter(rsa)
    formatter.SetHashAlgorithm("MD5")
    Dim hash As Byte() = ComputeStringHash(value)
    Return System.Convert.ToBase64String(formatter.CreateSignature(hash))
  End Function

  ''' <summary>
  ''' 校验数字签名
  ''' </summary>
  ''' <param name="publicKey">公钥</param>
  ''' <param name="value">签名的内容</param>
  ''' <param name="signature">数字签名</param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function VerifySignature(ByVal publicKey As String, ByVal value As String, ByVal signature As String) As Boolean
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    rsa.ImportCspBlob(System.Convert.FromBase64String(publicKey))
    Dim deformatter As New System.Security.Cryptography.RSAPKCS1SignatureDeformatter(rsa)
    deformatter.SetHashAlgorithm("MD5")
    Dim hash As Byte() = ComputeStringHash(value)
    Return deformatter.VerifySignature(hash, System.Convert.FromBase64String(signature))
  End Function

End Class
