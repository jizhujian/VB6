''' <summary>
''' 网卡函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(NetworkInterface.ClassId, NetworkInterface.InterfaceId, NetworkInterface.EventsId)> _
Public Class NetworkInterface

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "cd57949c-bfe6-4cda-a873-eb523bf89776"
  Public Const InterfaceId As String = "955cd993-30c1-4a63-8785-95a9c39ac012"
  Public Const EventsId As String = "9d64816b-18cd-4669-b93b-e96e05ccbfb2"

  Public Sub New()
    MyBase.New()
  End Sub

  ''' <summary>
  ''' 获取网卡MAC地址
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetEthernetMACAddresses() As String()
    Dim NetworkInterfaces As System.Net.NetworkInformation.NetworkInterface() = _
      System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
    Dim physicalAddresses As New System.Collections.ArrayList
    For Each NetworkInterface As System.Net.NetworkInformation.NetworkInterface In NetworkInterfaces
      If NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Ethernet OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.Ethernet3Megabit OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.FastEthernetFx OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.FastEthernetT OrElse _
         NetworkInterface.NetworkInterfaceType = System.Net.NetworkInformation.NetworkInterfaceType.GigabitEthernet Then
        physicalAddresses.Add(NetworkInterface.GetPhysicalAddress.ToString)
      End If
    Next
    Return physicalAddresses.ToArray(GetType(String))
  End Function

End Class
