<Microsoft.VisualBasic.ComClass(LicencesClient.ClassId, LicencesClient.InterfaceId, LicencesClient.EventsId)> _
Public Class LicencesClient

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = ""
  Public Const InterfaceId As String = ""
  Public Const EventsId As String = ""

  Public Sub New()
    MyBase.New()
  End Sub

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Private Function GetRegistryKey() As Microsoft.Win32.RegistryKey
    Dim RegistryKey1 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE").OpenSubKey("Jizhujian")
    If RegistryKey1 Is Nothing Then
      Return Nothing
    End If
    Dim RegistryKey2 As Microsoft.Win32.RegistryKey = RegistryKey1.OpenSubKey("AzK3")
    RegistryKey1.Close()
    Return RegistryKey2
  End Function

  '*****************************************************************************************
  '
  '*****************************************************************************************
  Public Function GetUserName() As String
    Dim RegistryKey As Microsoft.Win32.RegistryKey = GetRegistryKey()
    If RegistryKey Is Nothing Then
      Return String.Empty
    End If
    Dim UserName As String = TryCast(RegistryKey.GetValue("UserName", String.Empty), String)
    RegistryKey.Close()
    Return UserName
  End Function

  '*****************************************************************************************
  ' 根据UserName和Net MAC Address生成HASH值
  '*****************************************************************************************
  Public Function GetSerialNo() As Byte()

    Dim SerialNo As String = String.Empty

    Dim NetworkInterfaces As System.Net.NetworkInformation.NetworkInterface() = _
      System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
    Dim NetworkInterfacePhysicalAddresses As New System.Collections.ArrayList
    For Each NetworkInterface As System.Net.NetworkInformation.NetworkInterface In NetworkInterfaces
      If NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Ethernet OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Ethernet3Megabit OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.FastEthernetFx OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.FastEthernetT OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.GigabitEthernet OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Wireless80211 Then
        NetworkInterfacePhysicalAddresses.Add(NetworkInterface.GetPhysicalAddress.ToString)
      End If
    Next

    NetworkInterfacePhysicalAddresses.Sort()

    For Each NetworkInterfacePhysicalAddress As String In NetworkInterfacePhysicalAddresses
      SerialNo &= Microsoft.VisualBasic.ControlChars.Tab & NetworkInterfacePhysicalAddress
    Next

    Return (New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(System.Text.Encoding.Unicode.GetBytes(SerialNo))

  End Function

  '*****************************************************************************************
  ' 根据UserName+SerialNo＋Product＋ExpiredDate进行不对称解密
  '*****************************************************************************************
  Public Function ValidateProductRegister(ByVal Product As String) As Integer

    Dim SerialNoBytes As Byte() = GetSerialNo()
    Dim SerialNo As String = System.BitConverter.ToString(SerialNoBytes).Replace("-", "")

    Dim RegistryKey As Microsoft.Win32.RegistryKey = GetRegistryKey()
    If RegistryKey Is Nothing Then
      Return 0
    End If

    Dim UserName As String = GetUserName()
    If UserName.Length = 0 Then
      RegistryKey.Close()
      Return 0
    End If

    Dim ProductRegister As Byte() = TryCast(RegistryKey.GetValue(Product), Byte())
    If ProductRegister Is Nothing OrElse ProductRegister.Length = 0 Then
      RegistryKey.Close()
      Return 0
    End If

    RegistryKey.Close()

    Dim ProductRegisterSignature(ProductRegister.Length - 4 - 1) As Byte
    Dim ExpiredDateBytes(4 - 1) As Byte
    System.Array.Copy(ProductRegister, ProductRegisterSignature, ProductRegister.Length - 4)
    System.Array.Copy(ProductRegister, ProductRegister.Length - 4, ExpiredDateBytes, 0, 4)

    Dim UserNameBytes As Byte() = System.Text.Encoding.Unicode.GetBytes(UserName)
    Dim ProductBytes As Byte() = System.Text.Encoding.Unicode.GetBytes(Product)
    Dim ProductRegisterBytes(UserNameBytes.Length + SerialNoBytes.Length + ProductBytes.Length + ExpiredDateBytes.Length - 1) As Byte
    System.Array.Copy(UserNameBytes, ProductRegisterBytes, UserNameBytes.Length)
    System.Array.Copy(SerialNoBytes, 0, ProductRegisterBytes, UserNameBytes.Length, SerialNoBytes.Length)
    System.Array.Copy(ProductBytes, 0, ProductRegisterBytes, UserNameBytes.Length + SerialNoBytes.Length, ProductBytes.Length)
    System.Array.Copy(ExpiredDateBytes, 0, ProductRegisterBytes, UserNameBytes.Length + SerialNoBytes.Length + ProductBytes.Length, ExpiredDateBytes.Length)
    Dim ProductRegisterHash As Byte() = (New System.Security.Cryptography.MD5CryptoServiceProvider).ComputeHash(ProductRegisterBytes)

    Dim PublicKey As Byte() = { _
      6, 2, 0, 0, 0, 164, 0, 0, 82, 83, 65, 49, 0, 4, 0, 0, _
      1, 0, 1, 0, 125, 182, 7, 63, 187, 175, 69, 243, 36, 176, 97, 6, _
      89, 129, 222, 252, 72, 240, 210, 253, 204, 53, 245, 224, 186, 94, 87, 180, _
      79, 71, 33, 102, 53, 250, 7, 33, 241, 24, 73, 225, 207, 6, 11, 0, _
      100, 23, 128, 85, 124, 91, 222, 151, 56, 135, 63, 116, 11, 77, 4, 234, _
      23, 59, 152, 29, 171, 181, 231, 15, 241, 142, 135, 158, 134, 215, 118, 163, _
      44, 28, 188, 179, 176, 166, 149, 2, 159, 250, 41, 190, 197, 123, 245, 204, _
      97, 170, 133, 178, 46, 33, 109, 87, 245, 233, 198, 47, 204, 13, 86, 158, _
      178, 45, 81, 76, 156, 204, 142, 124, 94, 73, 89, 37, 54, 115, 6, 54, _
      139, 187, 129, 208}
    Dim rsa As New System.Security.Cryptography.RSACryptoServiceProvider
    rsa.ImportCspBlob(PublicKey)
    Dim RSADeformatter As New System.Security.Cryptography.RSAPKCS1SignatureDeformatter(rsa)
    RSADeformatter.SetHashAlgorithm("MD5")
    If Not RSADeformatter.VerifySignature(ProductRegisterHash, ProductRegisterSignature) Then
      Return 0
    End If

    Return System.BitConverter.ToInt32(ExpiredDateBytes, 0)

  End Function

End Class
