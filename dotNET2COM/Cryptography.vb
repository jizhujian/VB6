''' <summary>
''' ���ܺ�����
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Cryptography.ClassId, Cryptography.InterfaceId, Cryptography.EventsId)> _
Public Class Cryptography

  ''' <summary>
  ''' COMע�����
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "9e3c3a84-22cf-4bcd-b285-a712cb1eb94f"
  Public Const InterfaceId As String = "15285176-7117-4a5f-8942-0500f4f9a550"
  Public Const EventsId As String = "f4fdeb34-33b4-46d6-940d-5f59513116f9"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' �����ֽ������MD5��ϣֵ
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ComputeByteArrayHash(ByRef Value() As Byte) As Byte()
    Return (New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(Value)
  End Function

  ''' <summary>
  ''' �����ַ�����MD5��ϣֵ
  ''' </summary>
  ''' <param name="Value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ComputeStringHash(ByVal Value As String, Optional ByVal textCoding As File.TextEncoding = File.TextEncoding.Default) As Byte()
    Return ComputeByteArrayHash(File.GetTextEncoding(textCoding).GetBytes(Value))
  End Function

  ''' <summary>
  ''' �����ַ���
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
  ''' �����ַ���
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
  ''' ����˽Կ
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePrivateKey() As String
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    Return System.Convert.ToBase64String(rsa.ExportCspBlob(True))
  End Function

  ''' <summary>
  ''' ��ȡ��Կ
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
  ''' ��������ǩ��
  ''' </summary>
  ''' <param name="privateKey">˽Կ</param>
  ''' <param name="value">��Ҫǩ��������</param>
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
  ''' У������ǩ��
  ''' </summary>
  ''' <param name="publicKey">��Կ</param>
  ''' <param name="value">ǩ��������</param>
  ''' <param name="signature">����ǩ��</param>
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
