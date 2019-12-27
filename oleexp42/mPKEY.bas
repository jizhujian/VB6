Attribute VB_Name = "mPKEY"
Option Explicit
'propkey.bas
'Contains all entries from propkey.h as published in the Windows Platform SDK from Windows 10
'Converted by fafalone for VBForums

Private Sub DEFINE_PROPERTYKEY(Name As PROPERTYKEY, L As Long, w1 As Integer, w2 As Integer, B0 As Byte, b1 As Byte, b2 As Byte, B3 As Byte, b4 As Byte, b5 As Byte, b6 As Byte, b7 As Byte, pid As Long)
  With Name.fmtid
    .Data1 = L
    .Data2 = w1
    .Data3 = w2
    .Data4(0) = B0
    .Data4(1) = b1
    .Data4(2) = b2
    .Data4(3) = B3
    .Data4(4) = b4
    .Data4(5) = b5
    .Data4(6) = b6
    .Data4(7) = b7
  End With
  Name.pid = pid
End Sub

Public Function PKEY_Audio_ChannelCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 7)
PKEY_Audio_ChannelCount = pkk
End Function
Public Function PKEY_Audio_Compression() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 10)
PKEY_Audio_Compression = pkk
End Function
Public Function PKEY_Audio_EncodingBitrate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 4)
PKEY_Audio_EncodingBitrate = pkk
End Function
Public Function PKEY_Audio_Format() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 2)
PKEY_Audio_Format = pkk
End Function
Public Function PKEY_Audio_IsVariableBitRate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6822FEE, &H8C17, &H4D62, &H82, &H3C, &H8E, &H9C, &HFC, &HBD, &H1D, &H5C, 100)
PKEY_Audio_IsVariableBitRate = pkk
End Function
Public Function PKEY_Audio_PeakValue() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2579E5D0, &H1116, &H4084, &HBD, &H9A, &H9B, &H4F, &H7C, &HB4, &HDF, &H5E, 100)
PKEY_Audio_PeakValue = pkk
End Function
Public Function PKEY_Audio_SampleRate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 5)
PKEY_Audio_SampleRate = pkk
End Function
Public Function PKEY_Audio_SampleSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 6)
PKEY_Audio_SampleSize = pkk
End Function
Public Function PKEY_Audio_StreamName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 9)
PKEY_Audio_StreamName = pkk
End Function
Public Function PKEY_Audio_StreamNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 8)
PKEY_Audio_StreamNumber = pkk
End Function
Public Function PKEY_Calendar_Duration() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H293CA35A, &H9AA, &H4DD2, &HB1, &H80, &H1F, &HE2, &H45, &H72, &H8A, &H52, 100)
PKEY_Calendar_Duration = pkk
End Function
Public Function PKEY_Calendar_IsOnline() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBFEE9149, &HE3E2, &H49A7, &HA8, &H62, &HC0, &H59, &H88, &H14, &H5C, &HEC, 100)
PKEY_Calendar_IsOnline = pkk
End Function
Public Function PKEY_Calendar_IsRecurring() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H315B9C8D, &H80A9, &H4EF9, &HAE, &H16, &H8E, &H74, &H6D, &HA5, &H1D, &H70, 100)
PKEY_Calendar_IsRecurring = pkk
End Function
Public Function PKEY_Calendar_Location() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF6272D18, &HCECC, &H40B1, &HB2, &H6A, &H39, &H11, &H71, &H7A, &HA7, &HBD, 100)
PKEY_Calendar_Location = pkk
End Function
Public Function PKEY_Calendar_OptionalAttendeeAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD55BAE5A, &H3892, &H417A, &HA6, &H49, &HC6, &HAC, &H5A, &HAA, &HEA, &HB3, 100)
PKEY_Calendar_OptionalAttendeeAddresses = pkk
End Function
Public Function PKEY_Calendar_OptionalAttendeeNames() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9429607, &H582D, &H437F, &H84, &HC3, &HDE, &H93, &HA2, &HB2, &H4C, &H3C, 100)
PKEY_Calendar_OptionalAttendeeNames = pkk
End Function
Public Function PKEY_Calendar_OrganizerAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H744C8242, &H4DF5, &H456C, &HAB, &H9E, &H1, &H4E, &HFB, &H90, &H21, &HE3, 100)
PKEY_Calendar_OrganizerAddress = pkk
End Function
Public Function PKEY_Calendar_OrganizerName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAAA660F9, &H9865, &H458E, &HB4, &H84, &H1, &HBC, &H7F, &HE3, &H97, &H3E, 100)
PKEY_Calendar_OrganizerName = pkk
End Function
Public Function PKEY_Calendar_ReminderTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H72FC5BA4, &H24F9, &H4011, &H9F, &H3F, &HAD, &HD2, &H7A, &HFA, &HD8, &H18, 100)
PKEY_Calendar_ReminderTime = pkk
End Function
Public Function PKEY_Calendar_RequiredAttendeeAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBA7D6C3, &H568D, &H4159, &HAB, &H91, &H78, &H1A, &H91, &HFB, &H71, &HE5, 100)
PKEY_Calendar_RequiredAttendeeAddresses = pkk
End Function
Public Function PKEY_Calendar_RequiredAttendeeNames() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB33AF30B, &HF552, &H4584, &H93, &H6C, &HCB, &H93, &HE5, &HCD, &HA2, &H9F, 100)
PKEY_Calendar_RequiredAttendeeNames = pkk
End Function
Public Function PKEY_Calendar_Resources() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF58A38, &HC54B, &H4C40, &H86, &H96, &H97, &H23, &H59, &H80, &HEA, &HE1, 100)
PKEY_Calendar_Resources = pkk
End Function
Public Function PKEY_Calendar_ResponseStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H188C1F91, &H3C40, &H4132, &H9E, &HC5, &HD8, &HB0, &H3B, &H72, &HA8, &HA2, 100)
PKEY_Calendar_ResponseStatus = pkk
End Function
Public Function PKEY_Calendar_ShowTimeAs() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5BF396D4, &H5EB2, &H466F, &HBD, &HE9, &H2F, &HB3, &HF2, &H36, &H1D, &H6E, 100)
PKEY_Calendar_ShowTimeAs = pkk
End Function
Public Function PKEY_Calendar_ShowTimeAsText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H53DA57CF, &H62C0, &H45C4, &H81, &HDE, &H76, &H10, &HBC, &HEF, &HD7, &HF5, 100)
PKEY_Calendar_ShowTimeAsText = pkk
End Function
Public Function PKEY_Communication_AccountName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 9)
PKEY_Communication_AccountName = pkk
End Function
Public Function PKEY_Communication_DateItemExpires() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H428040AC, &HA177, &H4C8A, &H97, &H60, &HF6, &HF7, &H61, &H22, &H7F, &H9A, 100)
PKEY_Communication_DateItemExpires = pkk
End Function
Public Function PKEY_Communication_Direction() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8E531030, &HB960, &H4346, &HAE, &HD, &H66, &HBC, &H9A, &H86, &HFB, &H94, 100)
PKEY_Communication_Direction = pkk
End Function
Public Function PKEY_Communication_FollowupIconIndex() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H83A6347E, &H6FE4, &H4F40, &HBA, &H9C, &HC4, &H86, &H52, &H40, &HD1, &HF4, 100)
PKEY_Communication_FollowupIconIndex = pkk
End Function
Public Function PKEY_Communication_HeaderItem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9C34F84, &H2241, &H4401, &HB6, &H7, &HBD, &H20, &HED, &H75, &HAE, &H7F, 100)
PKEY_Communication_HeaderItem = pkk
End Function
Public Function PKEY_Communication_PolicyTag() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEC0B4191, &HAB0B, &H4C66, &H90, &HB6, &HC6, &H63, &H7C, &HDE, &HBB, &HAB, 100)
PKEY_Communication_PolicyTag = pkk
End Function
Public Function PKEY_Communication_SecurityFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8619A4B6, &H9F4D, &H4429, &H8C, &HF, &HB9, &H96, &HCA, &H59, &HE3, &H35, 100)
PKEY_Communication_SecurityFlags = pkk
End Function
Public Function PKEY_Communication_Suffix() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H807B653A, &H9E91, &H43EF, &H8F, &H97, &H11, &HCE, &H4, &HEE, &H20, &HC5, 100)
PKEY_Communication_Suffix = pkk
End Function
Public Function PKEY_Communication_TaskStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBE1A72C6, &H9A1D, &H46B7, &HAF, &HE7, &HAF, &HAF, &H8C, &HEF, &H49, &H99, 100)
PKEY_Communication_TaskStatus = pkk
End Function
Public Function PKEY_Communication_TaskStatusText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA6744477, &HC237, &H475B, &HA0, &H75, &H54, &HF3, &H44, &H98, &H29, &H2A, 100)
PKEY_Communication_TaskStatusText = pkk
End Function
Public Function PKEY_Computer_DecoratedFreeSpace() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 7)
PKEY_Computer_DecoratedFreeSpace = pkk
End Function
Public Function PKEY_Contact_AccountPictureDynamicVideo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB8BB018, &H2725, &H4B44, &H92, &HBA, &H79, &H33, &HAE, &HB2, &HDD, &HE7, 2)
PKEY_Contact_AccountPictureDynamicVideo = pkk
End Function
Public Function PKEY_Contact_AccountPictureLarge() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB8BB018, &H2725, &H4B44, &H92, &HBA, &H79, &H33, &HAE, &HB2, &HDD, &HE7, 3)
PKEY_Contact_AccountPictureLarge = pkk
End Function
Public Function PKEY_Contact_AccountPictureSmall() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB8BB018, &H2725, &H4B44, &H92, &HBA, &H79, &H33, &HAE, &HB2, &HDD, &HE7, 4)
PKEY_Contact_AccountPictureSmall = pkk
End Function
Public Function PKEY_Contact_Anniversary() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9AD5BADB, &HCEA7, &H4470, &HA0, &H3D, &HB8, &H4E, &H51, &HB9, &H94, &H9E, 100)
PKEY_Contact_Anniversary = pkk
End Function
Public Function PKEY_Contact_AssistantName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCD102C9C, &H5540, &H4A88, &HA6, &HF6, &H64, &HE4, &H98, &H1C, &H8C, &HD1, 100)
PKEY_Contact_AssistantName = pkk
End Function
Public Function PKEY_Contact_AssistantTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9A93244D, &HA7AD, &H4FF8, &H9B, &H99, &H45, &HEE, &H4C, &HC0, &H9A, &HF6, 100)
PKEY_Contact_AssistantTelephone = pkk
End Function
Public Function PKEY_Contact_Birthday() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 47)
PKEY_Contact_Birthday = pkk
End Function
Public Function PKEY_Contact_BusinessAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H730FB6DD, &HCF7C, &H426B, &HA0, &H3F, &HBD, &H16, &H6C, &HC9, &HEE, &H24, 100)
PKEY_Contact_BusinessAddress = pkk
End Function
Public Function PKEY_Contact_BusinessAddress1Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 119)
PKEY_Contact_BusinessAddress1Country = pkk
End Function
Public Function PKEY_Contact_BusinessAddress1Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 117)
PKEY_Contact_BusinessAddress1Locality = pkk
End Function
Public Function PKEY_Contact_BusinessAddress1PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 120)
PKEY_Contact_BusinessAddress1PostalCode = pkk
End Function
Public Function PKEY_Contact_BusinessAddress1Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 118)
PKEY_Contact_BusinessAddress1Region = pkk
End Function
Public Function PKEY_Contact_BusinessAddress1Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 116)
PKEY_Contact_BusinessAddress1Street = pkk
End Function
Public Function PKEY_Contact_BusinessAddress2Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 124)
PKEY_Contact_BusinessAddress2Country = pkk
End Function
Public Function PKEY_Contact_BusinessAddress2Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 122)
PKEY_Contact_BusinessAddress2Locality = pkk
End Function
Public Function PKEY_Contact_BusinessAddress2PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 125)
PKEY_Contact_BusinessAddress2PostalCode = pkk
End Function
Public Function PKEY_Contact_BusinessAddress2Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 123)
PKEY_Contact_BusinessAddress2Region = pkk
End Function
Public Function PKEY_Contact_BusinessAddress2Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 121)
PKEY_Contact_BusinessAddress2Street = pkk
End Function
Public Function PKEY_Contact_BusinessAddress3Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 129)
PKEY_Contact_BusinessAddress3Country = pkk
End Function
Public Function PKEY_Contact_BusinessAddress3Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 127)
PKEY_Contact_BusinessAddress3Locality = pkk
End Function
Public Function PKEY_Contact_BusinessAddress3PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 130)
PKEY_Contact_BusinessAddress3PostalCode = pkk
End Function
Public Function PKEY_Contact_BusinessAddress3Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 128)
PKEY_Contact_BusinessAddress3Region = pkk
End Function
Public Function PKEY_Contact_BusinessAddress3Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 126)
PKEY_Contact_BusinessAddress3Street = pkk
End Function
Public Function PKEY_Contact_BusinessAddressCity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H402B5934, &HEC5A, &H48C3, &H93, &HE6, &H85, &HE8, &H6A, &H2D, &H93, &H4E, 100)
PKEY_Contact_BusinessAddressCity = pkk
End Function
Public Function PKEY_Contact_BusinessAddressCountry() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB0B87314, &HFCF6, &H4FEB, &H8D, &HFF, &HA5, &HD, &HA6, &HAF, &H56, &H1C, 100)
PKEY_Contact_BusinessAddressCountry = pkk
End Function
Public Function PKEY_Contact_BusinessAddressPostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE1D4A09E, &HD758, &H4CD1, &HB6, &HEC, &H34, &HA8, &HB5, &HA7, &H3F, &H80, 100)
PKEY_Contact_BusinessAddressPostalCode = pkk
End Function
Public Function PKEY_Contact_BusinessAddressPostOfficeBox() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBC4E71CE, &H17F9, &H48D5, &HBE, &HE9, &H2, &H1D, &HF0, &HEA, &H54, &H9, 100)
PKEY_Contact_BusinessAddressPostOfficeBox = pkk
End Function
Public Function PKEY_Contact_BusinessAddressState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H446F787F, &H10C4, &H41CB, &HA6, &HC4, &H4D, &H3, &H43, &H55, &H15, &H97, 100)
PKEY_Contact_BusinessAddressState = pkk
End Function
Public Function PKEY_Contact_BusinessAddressStreet() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDDD1460F, &HC0BF, &H4553, &H8C, &HE4, &H10, &H43, &H3C, &H90, &H8F, &HB0, 100)
PKEY_Contact_BusinessAddressStreet = pkk
End Function
Public Function PKEY_Contact_BusinessEmailAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF271C659, &H7E5E, &H471F, &HBA, &H25, &H7F, &H77, &HB2, &H86, &HF8, &H36, 100)
PKEY_Contact_BusinessEmailAddresses = pkk
End Function
Public Function PKEY_Contact_BusinessFaxNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H91EFF6F3, &H2E27, &H42CA, &H93, &H3E, &H7C, &H99, &H9F, &HBE, &H31, &HB, 100)
PKEY_Contact_BusinessFaxNumber = pkk
End Function
Public Function PKEY_Contact_BusinessHomePage() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56310920, &H2491, &H4919, &H99, &HCE, &HEA, &HDB, &H6, &HFA, &HFD, &HB2, 100)
PKEY_Contact_BusinessHomePage = pkk
End Function
Public Function PKEY_Contact_BusinessTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6A15E5A0, &HA1E, &H4CD7, &HBB, &H8C, &HD2, &HF1, &HB0, &HC9, &H29, &HBC, 100)
PKEY_Contact_BusinessTelephone = pkk
End Function
Public Function PKEY_Contact_CallbackTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF53D1C3, &H49E0, &H4F7F, &H85, &H67, &H5A, &H82, &H1D, &H8A, &HC5, &H42, 100)
PKEY_Contact_CallbackTelephone = pkk
End Function
Public Function PKEY_Contact_CarTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8FDC6DEA, &HB929, &H412B, &HBA, &H90, &H39, &H7A, &H25, &H74, &H65, &HFE, 100)
PKEY_Contact_CarTelephone = pkk
End Function
Public Function PKEY_Contact_Children() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD4729704, &H8EF1, &H43EF, &H90, &H24, &H2B, &HD3, &H81, &H18, &H7F, &HD5, 100)
PKEY_Contact_Children = pkk
End Function
Public Function PKEY_Contact_CompanyMainTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8589E481, &H6040, &H473D, &HB1, &H71, &H7F, &HA8, &H9C, &H27, &H8, &HED, 100)
PKEY_Contact_CompanyMainTelephone = pkk
End Function
Public Function PKEY_Contact_ConnectedServiceDisplayName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H39B77F4F, &HA104, &H4863, &HB3, &H95, &H2D, &HB2, &HAD, &H8F, &H7B, &HC1, 100)
PKEY_Contact_ConnectedServiceDisplayName = pkk
End Function
Public Function PKEY_Contact_ConnectedServiceIdentities() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80F41EB8, &HAFC4, &H4208, &HAA, &H5F, &HCC, &HE2, &H1A, &H62, &H72, &H81, 100)
PKEY_Contact_ConnectedServiceIdentities = pkk
End Function
Public Function PKEY_Contact_ConnectedServiceName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB5C84C9E, &H5927, &H46B5, &HA3, &HCC, &H93, &H3C, &H21, &HB7, &H84, &H69, 100)
PKEY_Contact_ConnectedServiceName = pkk
End Function
Public Function PKEY_Contact_ConnectedServiceSupportedActions() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA19FB7A9, &H24B, &H4371, &HA8, &HBF, &H4D, &H29, &HC3, &HE4, &HE9, &HC9, 100)
PKEY_Contact_ConnectedServiceSupportedActions = pkk
End Function
Public Function PKEY_Contact_DataSuppliers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9660C283, &HFC3A, &H4A08, &HA0, &H96, &HEE, &HD3, &HAA, &HC4, &H6D, &HA2, 100)
PKEY_Contact_DataSuppliers = pkk
End Function
Public Function PKEY_Contact_Department() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFC9F7306, &HFF8F, &H4D49, &H9F, &HB6, &H3F, &HFE, &H5C, &H9, &H51, &HEC, 100)
PKEY_Contact_Department = pkk
End Function
Public Function PKEY_Contact_DisplayBusinessPhoneNumbers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H364028DA, &HD895, &H41FE, &HA5, &H84, &H30, &H2B, &H1B, &HB7, &HA, &H76, 100)
PKEY_Contact_DisplayBusinessPhoneNumbers = pkk
End Function
Public Function PKEY_Contact_DisplayHomePhoneNumbers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5068BCDF, &HD697, &H4D85, &H8C, &H53, &H1F, &H1C, &HDA, &HB0, &H17, &H63, 100)
PKEY_Contact_DisplayHomePhoneNumbers = pkk
End Function
Public Function PKEY_Contact_DisplayMobilePhoneNumbers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9CB0C358, &H9D7A, &H46B1, &HB4, &H66, &HDC, &HC6, &HF1, &HA3, &HD9, &H3D, 100)
PKEY_Contact_DisplayMobilePhoneNumbers = pkk
End Function
Public Function PKEY_Contact_DisplayOtherPhoneNumbers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3089873, &H8EE8, &H4191, &HBD, &H60, &HD3, &H1F, &H72, &HB7, &H90, &HB, 100)
PKEY_Contact_DisplayOtherPhoneNumbers = pkk
End Function
Public Function PKEY_Contact_EmailAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF8FA7FA3, &HD12B, &H4785, &H8A, &H4E, &H69, &H1A, &H94, &HF7, &HA3, &HE7, 100)
PKEY_Contact_EmailAddress = pkk
End Function
Public Function PKEY_Contact_EmailAddress2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H38965063, &HEDC8, &H4268, &H84, &H91, &HB7, &H72, &H31, &H72, &HCF, &H29, 100)
PKEY_Contact_EmailAddress2 = pkk
End Function
Public Function PKEY_Contact_EmailAddress3() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H644D37B4, &HE1B3, &H4BAD, &HB0, &H99, &H7E, &H7C, &H4, &H96, &H6A, &HCA, 100)
PKEY_Contact_EmailAddress3 = pkk
End Function
Public Function PKEY_Contact_EmailAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H84D8F337, &H981D, &H44B3, &H96, &H15, &HC7, &H59, &H6D, &HBA, &H17, &HE3, 100)
PKEY_Contact_EmailAddresses = pkk
End Function
Public Function PKEY_Contact_EmailName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCC6F4F24, &H6083, &H4BD4, &H87, &H54, &H67, &H4D, &HD, &HE8, &H7A, &HB8, 100)
PKEY_Contact_EmailName = pkk
End Function
Public Function PKEY_Contact_FileAsName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF1A24AA7, &H9CA7, &H40F6, &H89, &HEC, &H97, &HDE, &HF9, &HFF, &HE8, &HDB, 100)
PKEY_Contact_FileAsName = pkk
End Function
Public Function PKEY_Contact_FirstName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14977844, &H6B49, &H4AAD, &HA7, &H14, &HA4, &H51, &H3B, &HF6, &H4, &H60, 100)
PKEY_Contact_FirstName = pkk
End Function
Public Function PKEY_Contact_FullName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H635E9051, &H50A5, &H4BA2, &HB9, &HDB, &H4E, &HD0, &H56, &HC7, &H72, &H96, 100)
PKEY_Contact_FullName = pkk
End Function
Public Function PKEY_Contact_Gender() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3C8CEE58, &HD4F0, &H4CF9, &HB7, &H56, &H4E, &H5D, &H24, &H44, &H7B, &HCD, 100)
PKEY_Contact_Gender = pkk
End Function
Public Function PKEY_Contact_GenderValue() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3C8CEE58, &HD4F0, &H4CF9, &HB7, &H56, &H4E, &H5D, &H24, &H44, &H7B, &HCD, 101)
PKEY_Contact_GenderValue = pkk
End Function
Public Function PKEY_Contact_Hobbies() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5DC2253F, &H5E11, &H4ADF, &H9C, &HFE, &H91, &HD, &HD0, &H1E, &H3E, &H70, 100)
PKEY_Contact_Hobbies = pkk
End Function
Public Function PKEY_Contact_HomeAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H98F98354, &H617A, &H46B8, &H85, &H60, &H5B, &H1B, &H64, &HBF, &H1F, &H89, 100)
PKEY_Contact_HomeAddress = pkk
End Function
Public Function PKEY_Contact_HomeAddress1Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 104)
PKEY_Contact_HomeAddress1Country = pkk
End Function
Public Function PKEY_Contact_HomeAddress1Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 102)
PKEY_Contact_HomeAddress1Locality = pkk
End Function
Public Function PKEY_Contact_HomeAddress1PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 105)
PKEY_Contact_HomeAddress1PostalCode = pkk
End Function
Public Function PKEY_Contact_HomeAddress1Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 103)
PKEY_Contact_HomeAddress1Region = pkk
End Function
Public Function PKEY_Contact_HomeAddress1Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 101)
PKEY_Contact_HomeAddress1Street = pkk
End Function
Public Function PKEY_Contact_HomeAddress2Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 109)
PKEY_Contact_HomeAddress2Country = pkk
End Function
Public Function PKEY_Contact_HomeAddress2Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 107)
PKEY_Contact_HomeAddress2Locality = pkk
End Function
Public Function PKEY_Contact_HomeAddress2PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 110)
PKEY_Contact_HomeAddress2PostalCode = pkk
End Function
Public Function PKEY_Contact_HomeAddress2Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 108)
PKEY_Contact_HomeAddress2Region = pkk
End Function
Public Function PKEY_Contact_HomeAddress2Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 106)
PKEY_Contact_HomeAddress2Street = pkk
End Function
Public Function PKEY_Contact_HomeAddress3Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 114)
PKEY_Contact_HomeAddress3Country = pkk
End Function
Public Function PKEY_Contact_HomeAddress3Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 112)
PKEY_Contact_HomeAddress3Locality = pkk
End Function
Public Function PKEY_Contact_HomeAddress3PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 115)
PKEY_Contact_HomeAddress3PostalCode = pkk
End Function
Public Function PKEY_Contact_HomeAddress3Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 113)
PKEY_Contact_HomeAddress3Region = pkk
End Function
Public Function PKEY_Contact_HomeAddress3Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 111)
PKEY_Contact_HomeAddress3Street = pkk
End Function
Public Function PKEY_Contact_HomeAddressCity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 65)
PKEY_Contact_HomeAddressCity = pkk
End Function
Public Function PKEY_Contact_HomeAddressCountry() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8A65AA1, &HF4C9, &H43DD, &H9D, &HDF, &HA3, &H3D, &H8E, &H7E, &HAD, &H85, 100)
PKEY_Contact_HomeAddressCountry = pkk
End Function
Public Function PKEY_Contact_HomeAddressPostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8AFCC170, &H8A46, &H4B53, &H9E, &HEE, &H90, &HBA, &HE7, &H15, &H1E, &H62, 100)
PKEY_Contact_HomeAddressPostalCode = pkk
End Function
Public Function PKEY_Contact_HomeAddressPostOfficeBox() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7B9F6399, &HA3F, &H4B12, &H89, &HBD, &H4A, &HDC, &H51, &HC9, &H18, &HAF, 100)
PKEY_Contact_HomeAddressPostOfficeBox = pkk
End Function
Public Function PKEY_Contact_HomeAddressState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC89A23D0, &H7D6D, &H4EB8, &H87, &HD4, &H77, &H6A, &H82, &HD4, &H93, &HE5, 100)
PKEY_Contact_HomeAddressState = pkk
End Function
Public Function PKEY_Contact_HomeAddressStreet() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HADEF160, &HDB3F, &H4308, &H9A, &H21, &H6, &H23, &H7B, &H16, &HFA, &H2A, 100)
PKEY_Contact_HomeAddressStreet = pkk
End Function
Public Function PKEY_Contact_HomeEmailAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56C90E9D, &H9D46, &H4963, &H88, &H6F, &H2E, &H1C, &HD9, &HA6, &H94, &HEF, 100)
PKEY_Contact_HomeEmailAddresses = pkk
End Function
Public Function PKEY_Contact_HomeFaxNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H660E04D6, &H81AB, &H4977, &HA0, &H9F, &H82, &H31, &H31, &H13, &HAB, &H26, 100)
PKEY_Contact_HomeFaxNumber = pkk
End Function
Public Function PKEY_Contact_HomeTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 20)
PKEY_Contact_HomeTelephone = pkk
End Function
Public Function PKEY_Contact_IMAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD68DBD8A, &H3374, &H4B81, &H99, &H72, &H3E, &HC3, &H6, &H82, &HDB, &H3D, 100)
PKEY_Contact_IMAddress = pkk
End Function
Public Function PKEY_Contact_Initials() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF3D8F40D, &H50CB, &H44A2, &H97, &H18, &H40, &HCB, &H91, &H19, &H49, &H5D, 100)
PKEY_Contact_Initials = pkk
End Function
Public Function PKEY_Contact_JA_CompanyNamePhonetic() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H897B3694, &HFE9E, &H43E6, &H80, &H66, &H26, &HF, &H59, &HC, &H1, &H0, 2)
PKEY_Contact_JA_CompanyNamePhonetic = pkk
End Function
Public Function PKEY_Contact_JA_FirstNamePhonetic() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H897B3694, &HFE9E, &H43E6, &H80, &H66, &H26, &HF, &H59, &HC, &H1, &H0, 3)
PKEY_Contact_JA_FirstNamePhonetic = pkk
End Function
Public Function PKEY_Contact_JA_LastNamePhonetic() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H897B3694, &HFE9E, &H43E6, &H80, &H66, &H26, &HF, &H59, &HC, &H1, &H0, 4)
PKEY_Contact_JA_LastNamePhonetic = pkk
End Function
Public Function PKEY_Contact_JobInfo1CompanyAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 120)
PKEY_Contact_JobInfo1CompanyAddress = pkk
End Function
Public Function PKEY_Contact_JobInfo1CompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 102)
PKEY_Contact_JobInfo1CompanyName = pkk
End Function
Public Function PKEY_Contact_JobInfo1Department() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 106)
PKEY_Contact_JobInfo1Department = pkk
End Function
Public Function PKEY_Contact_JobInfo1Manager() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 105)
PKEY_Contact_JobInfo1Manager = pkk
End Function
Public Function PKEY_Contact_JobInfo1OfficeLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 104)
PKEY_Contact_JobInfo1OfficeLocation = pkk
End Function
Public Function PKEY_Contact_JobInfo1Title() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 103)
PKEY_Contact_JobInfo1Title = pkk
End Function
Public Function PKEY_Contact_JobInfo1YomiCompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 101)
PKEY_Contact_JobInfo1YomiCompanyName = pkk
End Function
Public Function PKEY_Contact_JobInfo2CompanyAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 121)
PKEY_Contact_JobInfo2CompanyAddress = pkk
End Function
Public Function PKEY_Contact_JobInfo2CompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 108)
PKEY_Contact_JobInfo2CompanyName = pkk
End Function
Public Function PKEY_Contact_JobInfo2Department() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 113)
PKEY_Contact_JobInfo2Department = pkk
End Function
Public Function PKEY_Contact_JobInfo2Manager() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 112)
PKEY_Contact_JobInfo2Manager = pkk
End Function
Public Function PKEY_Contact_JobInfo2OfficeLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 110)
PKEY_Contact_JobInfo2OfficeLocation = pkk
End Function
Public Function PKEY_Contact_JobInfo2Title() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 109)
PKEY_Contact_JobInfo2Title = pkk
End Function
Public Function PKEY_Contact_JobInfo2YomiCompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 107)
PKEY_Contact_JobInfo2YomiCompanyName = pkk
End Function
Public Function PKEY_Contact_JobInfo3CompanyAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 123)
PKEY_Contact_JobInfo3CompanyAddress = pkk
End Function
Public Function PKEY_Contact_JobInfo3CompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 115)
PKEY_Contact_JobInfo3CompanyName = pkk
End Function
Public Function PKEY_Contact_JobInfo3Department() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 119)
PKEY_Contact_JobInfo3Department = pkk
End Function
Public Function PKEY_Contact_JobInfo3Manager() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 118)
PKEY_Contact_JobInfo3Manager = pkk
End Function
Public Function PKEY_Contact_JobInfo3OfficeLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 117)
PKEY_Contact_JobInfo3OfficeLocation = pkk
End Function
Public Function PKEY_Contact_JobInfo3Title() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 116)
PKEY_Contact_JobInfo3Title = pkk
End Function
Public Function PKEY_Contact_JobInfo3YomiCompanyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 114)
PKEY_Contact_JobInfo3YomiCompanyName = pkk
End Function
Public Function PKEY_Contact_JobTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 6)
PKEY_Contact_JobTitle = pkk
End Function
Public Function PKEY_Contact_Label() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H97B0AD89, &HDF49, &H49CC, &H83, &H4E, &H66, &H9, &H74, &HFD, &H75, &H5B, 100)
PKEY_Contact_Label = pkk
End Function
Public Function PKEY_Contact_LastName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8F367200, &HC270, &H457C, &HB1, &HD4, &HE0, &H7C, &H5B, &HCD, &H90, &HC7, 100)
PKEY_Contact_LastName = pkk
End Function
Public Function PKEY_Contact_MailingAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC0AC206A, &H827E, &H4650, &H95, &HAE, &H77, &HE2, &HBB, &H74, &HFC, &HC9, 100)
PKEY_Contact_MailingAddress = pkk
End Function
Public Function PKEY_Contact_MiddleName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 71)
PKEY_Contact_MiddleName = pkk
End Function
Public Function PKEY_Contact_MobileTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 35)
PKEY_Contact_MobileTelephone = pkk
End Function
Public Function PKEY_Contact_NickName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 74)
PKEY_Contact_NickName = pkk
End Function
Public Function PKEY_Contact_OfficeLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 7)
PKEY_Contact_OfficeLocation = pkk
End Function
Public Function PKEY_Contact_OtherAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H508161FA, &H313B, &H43D5, &H83, &HA1, &HC1, &HAC, &HCF, &H68, &H62, &H2C, 100)
PKEY_Contact_OtherAddress = pkk
End Function
Public Function PKEY_Contact_OtherAddress1Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 134)
PKEY_Contact_OtherAddress1Country = pkk
End Function
Public Function PKEY_Contact_OtherAddress1Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 132)
PKEY_Contact_OtherAddress1Locality = pkk
End Function
Public Function PKEY_Contact_OtherAddress1PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 135)
PKEY_Contact_OtherAddress1PostalCode = pkk
End Function
Public Function PKEY_Contact_OtherAddress1Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 133)
PKEY_Contact_OtherAddress1Region = pkk
End Function
Public Function PKEY_Contact_OtherAddress1Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 131)
PKEY_Contact_OtherAddress1Street = pkk
End Function
Public Function PKEY_Contact_OtherAddress2Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 139)
PKEY_Contact_OtherAddress2Country = pkk
End Function
Public Function PKEY_Contact_OtherAddress2Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 137)
PKEY_Contact_OtherAddress2Locality = pkk
End Function
Public Function PKEY_Contact_OtherAddress2PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 140)
PKEY_Contact_OtherAddress2PostalCode = pkk
End Function
Public Function PKEY_Contact_OtherAddress2Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 138)
PKEY_Contact_OtherAddress2Region = pkk
End Function
Public Function PKEY_Contact_OtherAddress2Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 136)
PKEY_Contact_OtherAddress2Street = pkk
End Function
Public Function PKEY_Contact_OtherAddress3Country() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 144)
PKEY_Contact_OtherAddress3Country = pkk
End Function
Public Function PKEY_Contact_OtherAddress3Locality() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 142)
PKEY_Contact_OtherAddress3Locality = pkk
End Function
Public Function PKEY_Contact_OtherAddress3PostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 145)
PKEY_Contact_OtherAddress3PostalCode = pkk
End Function
Public Function PKEY_Contact_OtherAddress3Region() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 143)
PKEY_Contact_OtherAddress3Region = pkk
End Function
Public Function PKEY_Contact_OtherAddress3Street() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B6F596, &HD678, &H4BC1, &HB0, &H5F, &H2, &H3, &HD2, &H7E, &H8A, &HA1, 141)
PKEY_Contact_OtherAddress3Street = pkk
End Function
Public Function PKEY_Contact_OtherAddressCity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6E682923, &H7F7B, &H4F0C, &HA3, &H37, &HCF, &HCA, &H29, &H66, &H87, &HBF, 100)
PKEY_Contact_OtherAddressCity = pkk
End Function
Public Function PKEY_Contact_OtherAddressCountry() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8F167568, &HAAE, &H4322, &H8E, &HD9, &H60, &H55, &HB7, &HB0, &HE3, &H98, 100)
PKEY_Contact_OtherAddressCountry = pkk
End Function
Public Function PKEY_Contact_OtherAddressPostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95C656C1, &H2ABF, &H4148, &H9E, &HD3, &H9E, &HC6, &H2, &HE3, &HB7, &HCD, 100)
PKEY_Contact_OtherAddressPostalCode = pkk
End Function
Public Function PKEY_Contact_OtherAddressPostOfficeBox() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8B26EA41, &H58F, &H43F6, &HAE, &HCC, &H40, &H35, &H68, &H1C, &HE9, &H77, 100)
PKEY_Contact_OtherAddressPostOfficeBox = pkk
End Function
Public Function PKEY_Contact_OtherAddressState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H71B377D6, &HE570, &H425F, &HA1, &H70, &H80, &H9F, &HAE, &H73, &HE5, &H4E, 100)
PKEY_Contact_OtherAddressState = pkk
End Function
Public Function PKEY_Contact_OtherAddressStreet() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFF962609, &HB7D6, &H4999, &H86, &H2D, &H95, &H18, &HD, &H52, &H9A, &HEA, 100)
PKEY_Contact_OtherAddressStreet = pkk
End Function
Public Function PKEY_Contact_OtherEmailAddresses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H11D6336B, &H38C4, &H4EC9, &H84, &HD6, &HEB, &H38, &HD0, &HB1, &H50, &HAF, 100)
PKEY_Contact_OtherEmailAddresses = pkk
End Function
Public Function PKEY_Contact_PagerTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD6304E01, &HF8F5, &H4F45, &H8B, &H15, &HD0, &H24, &HA6, &H29, &H67, &H89, 100)
PKEY_Contact_PagerTelephone = pkk
End Function
Public Function PKEY_Contact_PersonalTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 69)
PKEY_Contact_PersonalTitle = pkk
End Function
Public Function PKEY_Contact_PhoneNumbersCanonical() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD042D2A1, &H927E, &H40B5, &HA5, &H3, &H6E, &HDB, &HD4, &H2A, &H51, &H7E, 100)
PKEY_Contact_PhoneNumbersCanonical = pkk
End Function
Public Function PKEY_Contact_Prefix() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 75)
PKEY_Contact_Prefix = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressCity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC8EA94F0, &HA9E3, &H4969, &HA9, &H4B, &H9C, &H62, &HA9, &H53, &H24, &HE0, 100)
PKEY_Contact_PrimaryAddressCity = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressCountry() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE53D799D, &HF3F, &H466E, &HB2, &HFF, &H74, &H63, &H4A, &H3C, &HB7, &HA4, 100)
PKEY_Contact_PrimaryAddressCountry = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressPostalCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H18BBD425, &HECFD, &H46EF, &HB6, &H12, &H7B, &H4A, &H60, &H34, &HED, &HA0, 100)
PKEY_Contact_PrimaryAddressPostalCode = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressPostOfficeBox() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDE5EF3C7, &H46E1, &H484E, &H99, &H99, &H62, &HC5, &H30, &H83, &H94, &HC1, 100)
PKEY_Contact_PrimaryAddressPostOfficeBox = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF1176DFE, &H7138, &H4640, &H8B, &H4C, &HAE, &H37, &H5D, &HC7, &HA, &H6D, 100)
PKEY_Contact_PrimaryAddressState = pkk
End Function
Public Function PKEY_Contact_PrimaryAddressStreet() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H63C25B20, &H96BE, &H488F, &H87, &H88, &HC0, &H9C, &H40, &H7A, &HD8, &H12, 100)
PKEY_Contact_PrimaryAddressStreet = pkk
End Function
Public Function PKEY_Contact_PrimaryEmailAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 48)
PKEY_Contact_PrimaryEmailAddress = pkk
End Function
Public Function PKEY_Contact_PrimaryTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 25)
PKEY_Contact_PrimaryTelephone = pkk
End Function
Public Function PKEY_Contact_Profession() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7268AF55, &H1CE4, &H4F6E, &HA4, &H1F, &HB6, &HE4, &HEF, &H10, &HE4, &HA9, 100)
PKEY_Contact_Profession = pkk
End Function
Public Function PKEY_Contact_SpouseName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9D2408B6, &H3167, &H422B, &H82, &HB0, &HF5, &H83, &HB7, &HA7, &HCF, &HE3, 100)
PKEY_Contact_SpouseName = pkk
End Function
Public Function PKEY_Contact_Suffix() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H176DC63C, &H2688, &H4E89, &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 73)
PKEY_Contact_Suffix = pkk
End Function
Public Function PKEY_Contact_TelexNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC554493C, &HC1F7, &H40C1, &HA7, &H6C, &HEF, &H8C, &H6, &H14, &H0, &H3E, 100)
PKEY_Contact_TelexNumber = pkk
End Function
Public Function PKEY_Contact_TTYTDDTelephone() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAAF16BAC, &H2B55, &H45E6, &H9F, &H6D, &H41, &H5E, &HB9, &H49, &H10, &HDF, 100)
PKEY_Contact_TTYTDDTelephone = pkk
End Function
Public Function PKEY_Contact_WebPage() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 18)
PKEY_Contact_WebPage = pkk
End Function
Public Function PKEY_Contact_Webpage2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 124)
PKEY_Contact_Webpage2 = pkk
End Function
Public Function PKEY_Contact_Webpage3() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF63DD8, &H22BD, &H4A5D, &HBA, &H34, &H5C, &HB0, &HB9, &HBD, &HCB, &H3, 125)
PKEY_Contact_Webpage3 = pkk
End Function
Public Function PKEY_AcquisitionID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H65A98875, &H3C80, &H40AB, &HAB, &HBC, &HEF, &HDA, &HF7, &H7D, &HBE, &HE2, 100)
PKEY_AcquisitionID = pkk
End Function
Public Function PKEY_ApplicationDefinedProperties() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCDBFC167, &H337E, &H41D8, &HAF, &H7C, &H8C, &H9, &H20, &H54, &H29, &HC7, 100)
PKEY_ApplicationDefinedProperties = pkk
End Function
Public Function PKEY_ApplicationName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 18)
PKEY_ApplicationName = pkk
End Function
Public Function PKEY_AppZoneIdentifier() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H502CFEAB, &H47EB, &H459C, &HB9, &H60, &HE6, &HD8, &H72, &H8F, &H77, &H1, 102)
PKEY_AppZoneIdentifier = pkk
End Function
Public Function PKEY_Author() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 4)
PKEY_Author = pkk
End Function
Public Function PKEY_CachedFileUpdaterContentIdForConflictResolution() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 114)
PKEY_CachedFileUpdaterContentIdForConflictResolution = pkk
End Function
Public Function PKEY_CachedFileUpdaterContentIdForStream() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 113)
PKEY_CachedFileUpdaterContentIdForStream = pkk
End Function
Public Function PKEY_Capacity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 3)
PKEY_Capacity = pkk
End Function
Public Function PKEY_Category() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 2)
PKEY_Category = pkk
End Function
Public Function PKEY_Comment() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 6)
PKEY_Comment = pkk
End Function
Public Function PKEY_Company() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 15)
PKEY_Company = pkk
End Function
Public Function PKEY_ComputerName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 5)
PKEY_ComputerName = pkk
End Function
Public Function PKEY_ContainedItems() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 29)
PKEY_ContainedItems = pkk
End Function
Public Function PKEY_ContentStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 27)
PKEY_ContentStatus = pkk
End Function
Public Function PKEY_ContentType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 26)
PKEY_ContentType = pkk
End Function
Public Function PKEY_Copyright() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 11)
PKEY_Copyright = pkk
End Function
Public Function PKEY_CreatorAppId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC2EA046E, &H33C, &H4E91, &HBD, &H5B, &HD4, &H94, &H2F, &H6B, &HBE, &H49, 2)
PKEY_CreatorAppId = pkk
End Function
Public Function PKEY_CreatorOpenWithUIOptions() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC2EA046E, &H33C, &H4E91, &HBD, &H5B, &HD4, &H94, &H2F, &H6B, &HBE, &H49, 3)
PKEY_CreatorOpenWithUIOptions = pkk
End Function
Public Function PKEY_DataObjectFormat() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1E81A3F8, &HA30F, &H4247, &HB9, &HEE, &H1D, &H3, &H68, &HA9, &H42, &H5C, 2)
PKEY_DataObjectFormat = pkk
End Function
Public Function PKEY_DateAccessed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 16)
PKEY_DateAccessed = pkk
End Function
Public Function PKEY_DateAcquired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2CBAA8F5, &HD81F, &H47CA, &HB1, &H7A, &HF8, &HD8, &H22, &H30, &H1, &H31, 100)
PKEY_DateAcquired = pkk
End Function
Public Function PKEY_DateArchived() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H43F8D7B7, &HA444, &H4F87, &H93, &H83, &H52, &H27, &H1C, &H9B, &H91, &H5C, 100)
PKEY_DateArchived = pkk
End Function
Public Function PKEY_DateCompleted() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H72FAB781, &HACDA, &H43E5, &HB1, &H55, &HB2, &H43, &H4F, &H85, &HE6, &H78, 100)
PKEY_DateCompleted = pkk
End Function
Public Function PKEY_DateCreated() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 15)
PKEY_DateCreated = pkk
End Function
Public Function PKEY_DateImported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 18258)
PKEY_DateImported = pkk
End Function
Public Function PKEY_DateModified() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 14)
PKEY_DateModified = pkk
End Function
Public Function PKEY_DefaultSaveLocationDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 10)
PKEY_DefaultSaveLocationDisplay = pkk
End Function
Public Function PKEY_DueDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3F8472B5, &HE0AF, &H4DB2, &H80, &H71, &HC5, &H3F, &HE7, &H6A, &HE7, &HCE, 100)
PKEY_DueDate = pkk
End Function
Public Function PKEY_EndDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC75FAA05, &H96FD, &H49E7, &H9C, &HB4, &H9F, &H60, &H10, &H82, &HD5, &H53, 100)
PKEY_EndDate = pkk
End Function
Public Function PKEY_ExpandoProperties() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6FA20DE6, &HD11C, &H4D9D, &HA1, &H54, &H64, &H31, &H76, &H28, &HC1, &H2D, 100)
PKEY_ExpandoProperties = pkk
End Function
Public Function PKEY_FileAllocationSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 18)
PKEY_FileAllocationSize = pkk
End Function
Public Function PKEY_FileAttributes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 13)
PKEY_FileAttributes = pkk
End Function
Public Function PKEY_FileCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 12)
PKEY_FileCount = pkk
End Function
Public Function PKEY_FileDescription() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 3)
PKEY_FileDescription = pkk
End Function
Public Function PKEY_FileExtension() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE4F10A3C, &H49E6, &H405D, &H82, &H88, &HA2, &H3B, &HD4, &HEE, &HAA, &H6C, 100)
PKEY_FileExtension = pkk
End Function
Public Function PKEY_FileFRN() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 21)
PKEY_FileFRN = pkk
End Function
Public Function PKEY_FileName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H41CF5AE0, &HF75A, &H4806, &HBD, &H87, &H59, &HC7, &HD9, &H24, &H8E, &HB9, 100)
PKEY_FileName = pkk
End Function
Public Function PKEY_FileOfflineAvailabilityStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 100)
PKEY_FileOfflineAvailabilityStatus = pkk
End Function
Public Function PKEY_FileOwner() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B34, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 4)
PKEY_FileOwner = pkk
End Function
Public Function PKEY_FilePlaceholderStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB2F9B9D6, &HFEC4, &H4DD5, &H94, &HD7, &H89, &H57, &H48, &H8C, &H80, &H7B, 2)
PKEY_FilePlaceholderStatus = pkk
End Function
Public Function PKEY_FileVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 4)
PKEY_FileVersion = pkk
End Function
Public Function PKEY_FindData() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 0)
PKEY_FindData = pkk
End Function
Public Function PKEY_FlagColor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H67DF94DE, &HCA7, &H4D6F, &HB7, &H92, &H5, &H3A, &H3E, &H4F, &H3, &HCF, 100)
PKEY_FlagColor = pkk
End Function
Public Function PKEY_FlagColorText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H45EAE747, &H8E2A, &H40AE, &H8C, &HBF, &HCA, &H52, &HAB, &HA6, &H15, &H2A, 100)
PKEY_FlagColorText = pkk
End Function
Public Function PKEY_FlagStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 12)
PKEY_FlagStatus = pkk
End Function
Public Function PKEY_FlagStatusText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDC54FD2E, &H189D, &H4871, &HAA, &H1, &H8, &HC2, &HF5, &H7A, &H4A, &HBC, 100)
PKEY_FlagStatusText = pkk
End Function
Public Function PKEY_FolderKind() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 101)
PKEY_FolderKind = pkk
End Function
Public Function PKEY_FolderNameDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 25)
PKEY_FolderNameDisplay = pkk
End Function
Public Function PKEY_FreeSpace() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 2)
PKEY_FreeSpace = pkk
End Function
Public Function PKEY_FullText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1E3EE840, &HBC2B, &H476C, &H82, &H37, &H2A, &HCD, &H1A, &H83, &H9B, &H22, 6)
PKEY_FullText = pkk
End Function
Public Function PKEY_HighKeywords() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 24)
PKEY_HighKeywords = pkk
End Function
Public Function PKEY_Identity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA26F4AFC, &H7346, &H4299, &HBE, &H47, &HEB, &H1A, &HE6, &H13, &H13, &H9F, 100)
PKEY_Identity = pkk
End Function
Public Function PKEY_Identity_Blob() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8C3B93A4, &HBAED, &H1A83, &H9A, &H32, &H10, &H2E, &HE3, &H13, &HF6, &HEB, 100)
PKEY_Identity_Blob = pkk
End Function
Public Function PKEY_Identity_DisplayName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7D683FC9, &HD155, &H45A8, &HBB, &H1F, &H89, &HD1, &H9B, &HCB, &H79, &H2F, 100)
PKEY_Identity_DisplayName = pkk
End Function
Public Function PKEY_Identity_InternetSid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D6D5D49, &H265D, &H4688, &H9F, &H4E, &H1F, &HDD, &H33, &HE7, &HCC, &H83, 100)
PKEY_Identity_InternetSid = pkk
End Function
Public Function PKEY_Identity_IsMeIdentity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA4108708, &H9DF, &H4377, &H9D, &HFC, &H6D, &H99, &H98, &H6D, &H5A, &H67, 100)
PKEY_Identity_IsMeIdentity = pkk
End Function
Public Function PKEY_Identity_KeyProviderContext() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA26F4AFC, &H7346, &H4299, &HBE, &H47, &HEB, &H1A, &HE6, &H13, &H13, &H9F, 17)
PKEY_Identity_KeyProviderContext = pkk
End Function
Public Function PKEY_Identity_KeyProviderName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA26F4AFC, &H7346, &H4299, &HBE, &H47, &HEB, &H1A, &HE6, &H13, &H13, &H9F, 16)
PKEY_Identity_KeyProviderName = pkk
End Function
Public Function PKEY_Identity_LogonStatusString() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF18DEDF3, &H337F, &H42C0, &H9E, &H3, &HCE, &HE0, &H87, &H8, &HA8, &HC3, 100)
PKEY_Identity_LogonStatusString = pkk
End Function
Public Function PKEY_Identity_PrimaryEmailAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCC16823, &HBAED, &H4F24, &H9B, &H32, &HA0, &H98, &H21, &H17, &HF7, &HFA, 100)
PKEY_Identity_PrimaryEmailAddress = pkk
End Function
Public Function PKEY_Identity_PrimarySid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2B1B801E, &HC0C1, &H4987, &H9E, &HC5, &H72, &HFA, &H89, &H81, &H47, &H87, 100)
PKEY_Identity_PrimarySid = pkk
End Function
Public Function PKEY_Identity_ProviderData() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8A74B92, &H361B, &H4E9A, &HB7, &H22, &H7C, &H4A, &H73, &H30, &HA3, &H12, 100)
PKEY_Identity_ProviderData = pkk
End Function
Public Function PKEY_Identity_ProviderID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H74A7DE49, &HFA11, &H4D3D, &HA0, &H6, &HDB, &H7E, &H8, &H67, &H59, &H16, 100)
PKEY_Identity_ProviderID = pkk
End Function
Public Function PKEY_Identity_QualifiedUserName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDA520E51, &HF4E9, &H4739, &HAC, &H82, &H2, &HE0, &HA9, &H5C, &H90, &H30, 100)
PKEY_Identity_QualifiedUserName = pkk
End Function
Public Function PKEY_Identity_UniqueID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE55FC3B0, &H2B60, &H4220, &H91, &H8E, &HB2, &H1E, &H8B, &HF1, &H60, &H16, 100)
PKEY_Identity_UniqueID = pkk
End Function
Public Function PKEY_Identity_UserName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC4322503, &H78CA, &H49C6, &H9A, &HCC, &HA6, &H8E, &H2A, &HFD, &H7B, &H6B, 100)
PKEY_Identity_UserName = pkk
End Function
Public Function PKEY_IdentityProvider_Name() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB96EFF7B, &H35CA, &H4A35, &H86, &H7, &H29, &HE3, &HA5, &H4C, &H46, &HEA, 100)
PKEY_IdentityProvider_Name = pkk
End Function
Public Function PKEY_IdentityProvider_Picture() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2425166F, &H5642, &H4864, &H99, &H2F, &H98, &HFD, &H98, &HF2, &H94, &HC3, 100)
PKEY_IdentityProvider_Picture = pkk
End Function
Public Function PKEY_ImageParsingName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD7750EE0, &HC6A4, &H48EC, &HB5, &H3E, &HB8, &H7B, &H52, &HE6, &HD0, &H73, 100)
PKEY_ImageParsingName = pkk
End Function
Public Function PKEY_Importance() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 11)
PKEY_Importance = pkk
End Function
Public Function PKEY_ImportanceText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA3B29791, &H7713, &H4E1D, &HBB, &H40, &H17, &HDB, &H85, &HF0, &H18, &H31, 100)
PKEY_ImportanceText = pkk
End Function
Public Function PKEY_IsAttachment() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF23F425C, &H71A1, &H4FA8, &H92, &H2F, &H67, &H8E, &HA4, &HA6, &H4, &H8, 100)
PKEY_IsAttachment = pkk
End Function
Public Function PKEY_IsDefaultNonOwnerSaveLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 5)
PKEY_IsDefaultNonOwnerSaveLocation = pkk
End Function
Public Function PKEY_IsDefaultSaveLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 3)
PKEY_IsDefaultSaveLocation = pkk
End Function
Public Function PKEY_IsDeleted() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5CDA5FC8, &H33EE, &H4FF3, &H90, &H94, &HAE, &H7B, &HD8, &H86, &H8C, &H4D, 100)
PKEY_IsDeleted = pkk
End Function
Public Function PKEY_IsEncrypted() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H90E5E14E, &H648B, &H4826, &HB2, &HAA, &HAC, &HAF, &H79, &HE, &H35, &H13, 10)
PKEY_IsEncrypted = pkk
End Function
Public Function PKEY_IsFlagged() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5DA84765, &HE3FF, &H4278, &H86, &HB0, &HA2, &H79, &H67, &HFB, &HDD, &H3, 100)
PKEY_IsFlagged = pkk
End Function
Public Function PKEY_IsFlaggedComplete() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA6F360D2, &H55F9, &H48DE, &HB9, &H9, &H62, &HE, &H9, &HA, &H64, &H7C, 100)
PKEY_IsFlaggedComplete = pkk
End Function
Public Function PKEY_IsIncomplete() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H346C8BD1, &H2E6A, &H4C45, &H89, &HA4, &H61, &HB7, &H8E, &H8E, &H70, &HF, 100)
PKEY_IsIncomplete = pkk
End Function
Public Function PKEY_IsLocationSupported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 8)
PKEY_IsLocationSupported = pkk
End Function
Public Function PKEY_IsPinnedToNameSpaceTree() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 2)
PKEY_IsPinnedToNameSpaceTree = pkk
End Function
Public Function PKEY_IsRead() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 10)
PKEY_IsRead = pkk
End Function
Public Function PKEY_IsSearchOnlyItem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 4)
PKEY_IsSearchOnlyItem = pkk
End Function
Public Function PKEY_IsSendToTarget() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 33)
PKEY_IsSendToTarget = pkk
End Function
Public Function PKEY_IsShared() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF884C5B, &H2BFE, &H41BB, &HAA, &HE5, &H76, &HEE, &HDF, &H4F, &H99, &H2, 100)
PKEY_IsShared = pkk
End Function
Public Function PKEY_ItemAuthors() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD0A04F0A, &H462A, &H48A4, &HBB, &H2F, &H37, &H6, &HE8, &H8D, &HBD, &H7D, 100)
PKEY_ItemAuthors = pkk
End Function
Public Function PKEY_ItemClassType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H48658AD, &H2DB8, &H41A4, &HBB, &HB6, &HAC, &H1E, &HF1, &H20, &H7E, &HB1, 100)
PKEY_ItemClassType = pkk
End Function
Public Function PKEY_ItemDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF7DB74B4, &H4287, &H4103, &HAF, &HBA, &HF1, &HB1, &H3D, &HCD, &H75, &HCF, 100)
PKEY_ItemDate = pkk
End Function
Public Function PKEY_ItemFolderNameDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 2)
PKEY_ItemFolderNameDisplay = pkk
End Function
Public Function PKEY_ItemFolderPathDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 6)
PKEY_ItemFolderPathDisplay = pkk
End Function
Public Function PKEY_ItemFolderPathDisplayNarrow() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDABD30ED, &H43, &H4789, &HA7, &HF8, &HD0, &H13, &HA4, &H73, &H66, &H22, 100)
PKEY_ItemFolderPathDisplayNarrow = pkk
End Function
Public Function PKEY_ItemName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6B8DA074, &H3B5C, &H43BC, &H88, &H6F, &HA, &H2C, &HDC, &HE0, &HB, &H6F, 100)
PKEY_ItemName = pkk
End Function
Public Function PKEY_ItemNameDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 10)
PKEY_ItemNameDisplay = pkk
End Function
Public Function PKEY_ItemNameDisplayWithoutExtension() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 24)
PKEY_ItemNameDisplayWithoutExtension = pkk
End Function
Public Function PKEY_ItemNamePrefix() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD7313FF1, &HA77A, &H401C, &H8C, &H99, &H3D, &HBD, &HD6, &H8A, &HDD, &H36, 100)
PKEY_ItemNamePrefix = pkk
End Function
Public Function PKEY_ItemNameSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 23)
PKEY_ItemNameSortOverride = pkk
End Function
Public Function PKEY_ItemParticipants() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD4D0AA16, &H9948, &H41A4, &HAA, &H85, &HD9, &H7F, &HF9, &H64, &H69, &H93, 100)
PKEY_ItemParticipants = pkk
End Function
Public Function PKEY_ItemPathDisplay() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 7)
PKEY_ItemPathDisplay = pkk
End Function
Public Function PKEY_ItemPathDisplayNarrow() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 8)
PKEY_ItemPathDisplayNarrow = pkk
End Function
Public Function PKEY_ItemSubType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 37)
PKEY_ItemSubType = pkk
End Function
Public Function PKEY_ItemType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 11)
PKEY_ItemType = pkk
End Function
Public Function PKEY_ItemTypeText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 4)
PKEY_ItemTypeText = pkk
End Function
Public Function PKEY_ItemUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49691C90, &H7E17, &H101A, &HA9, &H1C, &H8, &H0, &H2B, &H2E, &HCD, &HA9, 9)
PKEY_ItemUrl = pkk
End Function
Public Function PKEY_Keywords() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 5)
PKEY_Keywords = pkk
End Function
Public Function PKEY_Kind() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1E3EE840, &HBC2B, &H476C, &H82, &H37, &H2A, &HCD, &H1A, &H83, &H9B, &H22, 3)
PKEY_Kind = pkk
End Function
Public Function PKEY_KindText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF04BEF95, &HC585, &H4197, &HA2, &HB7, &HDF, &H46, &HFD, &HC9, &HEE, &H6D, 100)
PKEY_KindText = pkk
End Function
Public Function PKEY_Language() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 28)
PKEY_Language = pkk
End Function
Public Function PKEY_LastSyncError() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 107)
PKEY_LastSyncError = pkk
End Function
Public Function PKEY_LastWriterPackageFamilyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H502CFEAB, &H47EB, &H459C, &HB9, &H60, &HE6, &HD8, &H72, &H8F, &H77, &H1, 101)
PKEY_LastWriterPackageFamilyName = pkk
End Function
Public Function PKEY_LowKeywords() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 25)
PKEY_LowKeywords = pkk
End Function
Public Function PKEY_MediumKeywords() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 26)
PKEY_MediumKeywords = pkk
End Function
Public Function PKEY_MileageInformation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFDF84370, &H31A, &H4ADD, &H9E, &H91, &HD, &H77, &H5F, &H1C, &H66, &H5, 100)
PKEY_MileageInformation = pkk
End Function
Public Function PKEY_MIMEType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E350, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 5)
PKEY_MIMEType = pkk
End Function
Public Function PKEY_Null() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, 0)
PKEY_Null = pkk
End Function
Public Function PKEY_OfflineAvailability() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA94688B6, &H7D9F, &H4570, &HA6, &H48, &HE3, &HDF, &HC0, &HAB, &H2B, &H3F, 100)
PKEY_OfflineAvailability = pkk
End Function
Public Function PKEY_OfflineStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D24888F, &H4718, &H4BDA, &HAF, &HED, &HEA, &HF, &HB4, &H38, &H6C, &HD8, 100)
PKEY_OfflineStatus = pkk
End Function
Public Function PKEY_OriginalFileName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 6)
PKEY_OriginalFileName = pkk
End Function
Public Function PKEY_OwnerSID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5D76B67F, &H9B3D, &H44BB, &HB6, &HAE, &H25, &HDA, &H4F, &H63, &H8A, &H67, 6)
PKEY_OwnerSID = pkk
End Function
Public Function PKEY_ParentalRating() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 21)
PKEY_ParentalRating = pkk
End Function
Public Function PKEY_ParentalRatingReason() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10984E0A, &HF9F2, &H4321, &HB7, &HEF, &HBA, &HF1, &H95, &HAF, &H43, &H19, 100)
PKEY_ParentalRatingReason = pkk
End Function
Public Function PKEY_ParentalRatingsOrganization() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7FE0840, &H1344, &H46F0, &H8D, &H37, &H52, &HED, &H71, &H2A, &H4B, &HF9, 100)
PKEY_ParentalRatingsOrganization = pkk
End Function
Public Function PKEY_ParsingBindContext() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDFB9A04D, &H362F, &H4CA3, &HB3, &HB, &H2, &H54, &HB1, &H7B, &H5B, &H84, 100)
PKEY_ParsingBindContext = pkk
End Function
Public Function PKEY_ParsingName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 24)
PKEY_ParsingName = pkk
End Function
Public Function PKEY_ParsingPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 30)
PKEY_ParsingPath = pkk
End Function
Public Function PKEY_PerceivedType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 9)
PKEY_PerceivedType = pkk
End Function
Public Function PKEY_PercentFull() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 5)
PKEY_PercentFull = pkk
End Function
Public Function PKEY_Priority() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9C1FCF74, &H2D97, &H41BA, &HB4, &HAE, &HCB, &H2E, &H36, &H61, &HA6, &HE4, 5)
PKEY_Priority = pkk
End Function
Public Function PKEY_PriorityText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD98BE98B, &HB86B, &H4095, &HBF, &H52, &H9D, &H23, &HB2, &HE0, &HA7, &H52, 100)
PKEY_PriorityText = pkk
End Function
Public Function PKEY_Project() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H39A7F922, &H477C, &H48DE, &H8B, &HC8, &HB2, &H84, &H41, &HE3, &H42, &HE3, 100)
PKEY_Project = pkk
End Function
Public Function PKEY_ProviderItemID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF21D9941, &H81F0, &H471A, &HAD, &HEE, &H4E, &H74, &HB4, &H92, &H17, &HED, 100)
PKEY_ProviderItemID = pkk
End Function
Public Function PKEY_Rating() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 9)
PKEY_Rating = pkk
End Function
Public Function PKEY_RatingText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H90197CA7, &HFD8F, &H4E8C, &H9D, &HA3, &HB5, &H7E, &H1E, &H60, &H92, &H95, 100)
PKEY_RatingText = pkk
End Function
Public Function PKEY_RemoteConflictingFile() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 115)
PKEY_RemoteConflictingFile = pkk
End Function
Public Function PKEY_Security_EncryptionOwners() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDE621B8F, &HE125, &H43A3, &HA3, &H2D, &H56, &H65, &H44, &H6D, &H63, &H2A, 25)
PKEY_Security_EncryptionOwners = pkk
End Function
Public Function PKEY_Sensitivity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF8D3F6AC, &H4874, &H42CB, &HBE, &H59, &HAB, &H45, &H4B, &H30, &H71, &H6A, 100)
PKEY_Sensitivity = pkk
End Function
Public Function PKEY_SensitivityText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD0C7F054, &H3F72, &H4725, &H85, &H27, &H12, &H9A, &H57, &H7C, &HB2, &H69, 100)
PKEY_SensitivityText = pkk
End Function
Public Function PKEY_SFGAOFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 25)
PKEY_SFGAOFlags = pkk
End Function
Public Function PKEY_SharedWith() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF884C5B, &H2BFE, &H41BB, &HAA, &HE5, &H76, &HEE, &HDF, &H4F, &H99, &H2, 200)
PKEY_SharedWith = pkk
End Function
Public Function PKEY_ShareUserRating() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 12)
PKEY_ShareUserRating = pkk
End Function
Public Function PKEY_SharingStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF884C5B, &H2BFE, &H41BB, &HAA, &HE5, &H76, &HEE, &HDF, &H4F, &H99, &H2, 300)
PKEY_SharingStatus = pkk
End Function
Public Function PKEY_Shell_OmitFromView() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDE35258C, &HC695, &H4CBC, &HB9, &H82, &H38, &HB0, &HAD, &H24, &HCE, &HD0, 2)
PKEY_Shell_OmitFromView = pkk
End Function
Public Function PKEY_SimpleRating() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA09F084E, &HAD41, &H489F, &H80, &H76, &HAA, &H5B, &HE3, &H8, &H2B, &HCA, 100)
PKEY_SimpleRating = pkk
End Function
Public Function PKEY_Size() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 12)
PKEY_Size = pkk
End Function
Public Function PKEY_SoftwareUsed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 305)
PKEY_SoftwareUsed = pkk
End Function
Public Function PKEY_SourceItem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H668CDFA5, &H7A1B, &H4323, &HAE, &H4B, &HE5, &H27, &H39, &H3A, &H1D, &H81, 100)
PKEY_SourceItem = pkk
End Function
Public Function PKEY_SourcePackageFamilyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFFAE9DB7, &H1C8D, &H43FF, &H81, &H8C, &H84, &H40, &H3A, &HA3, &H73, &H2D, 100)
PKEY_SourcePackageFamilyName = pkk
End Function
Public Function PKEY_StartDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H48FD6EC8, &H8A12, &H4CDF, &HA0, &H3E, &H4E, &HC5, &HA5, &H11, &HED, &HDE, 100)
PKEY_StartDate = pkk
End Function
Public Function PKEY_Status() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H214A1, &H0, &H0, &HC0, &H0, &H0, &H0, &H0, &H0, &H0, &H46, 9)
PKEY_Status = pkk
End Function
Public Function PKEY_StorageProviderError() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 109)
PKEY_StorageProviderError = pkk
End Function
Public Function PKEY_StorageProviderFileChecksum() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB2F9B9D6, &HFEC4, &H4DD5, &H94, &HD7, &H89, &H57, &H48, &H8C, &H80, &H7B, 5)
PKEY_StorageProviderFileChecksum = pkk
End Function
Public Function PKEY_StorageProviderFileIdentifier() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB2F9B9D6, &HFEC4, &H4DD5, &H94, &HD7, &H89, &H57, &H48, &H8C, &H80, &H7B, 3)
PKEY_StorageProviderFileIdentifier = pkk
End Function
Public Function PKEY_StorageProviderFileRemoteUri() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 112)
PKEY_StorageProviderFileRemoteUri = pkk
End Function
Public Function PKEY_StorageProviderFileVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB2F9B9D6, &HFEC4, &H4DD5, &H94, &HD7, &H89, &H57, &H48, &H8C, &H80, &H7B, 4)
PKEY_StorageProviderFileVersion = pkk
End Function
Public Function PKEY_StorageProviderFileVersionWaterline() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB2F9B9D6, &HFEC4, &H4DD5, &H94, &HD7, &H89, &H57, &H48, &H8C, &H80, &H7B, 6)
PKEY_StorageProviderFileVersionWaterline = pkk
End Function
Public Function PKEY_StorageProviderId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 108)
PKEY_StorageProviderId = pkk
End Function
Public Function PKEY_StorageProviderShareStatuses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 111)
PKEY_StorageProviderShareStatuses = pkk
End Function
Public Function PKEY_StorageProviderSharingStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 117)
PKEY_StorageProviderSharingStatus = pkk
End Function
Public Function PKEY_StorageProviderStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 110)
PKEY_StorageProviderStatus = pkk
End Function
Public Function PKEY_Subject() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 3)
PKEY_Subject = pkk
End Function
Public Function PKEY_SyncTransferStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 103)
PKEY_SyncTransferStatus = pkk
End Function
Public Function PKEY_Thumbnail() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 17)
PKEY_Thumbnail = pkk
End Function
Public Function PKEY_ThumbnailCacheId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H446D16B1, &H8DAD, &H4870, &HA7, &H48, &H40, &H2E, &HA4, &H3D, &H78, &H8C, 100)
PKEY_ThumbnailCacheId = pkk
End Function
Public Function PKEY_ThumbnailStream() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 27)
PKEY_ThumbnailStream = pkk
End Function
Public Function PKEY_Title() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 2)
PKEY_Title = pkk
End Function
Public Function PKEY_TitleSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF0F7984D, &H222E, &H4AD2, &H82, &HAB, &H1D, &HD8, &HEA, &H40, &HE5, &H7E, 300)
PKEY_TitleSortOverride = pkk
End Function
Public Function PKEY_TotalFileSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 14)
PKEY_TotalFileSize = pkk
End Function
Public Function PKEY_Trademarks() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 9)
PKEY_Trademarks = pkk
End Function
Public Function PKEY_TransferOrder() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 106)
PKEY_TransferOrder = pkk
End Function
Public Function PKEY_TransferPosition() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 104)
PKEY_TransferPosition = pkk
End Function
Public Function PKEY_TransferSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCEFF153, &HE839, &H4CF3, &HA9, &HE7, &HEA, &H22, &H83, &H20, &H94, &HB8, 105)
PKEY_TransferSize = pkk
End Function
Public Function PKEY_VolumeId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H446D16B1, &H8DAD, &H4870, &HA7, &H48, &H40, &H2E, &HA4, &H3D, &H78, &H8C, 104)
PKEY_VolumeId = pkk
End Function
Public Function PKEY_ZoneIdentifier() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H502CFEAB, &H47EB, &H459C, &HB9, &H60, &HE6, &HD8, &H72, &H8F, &H77, &H1, 100)
PKEY_ZoneIdentifier = pkk
End Function
Public Function PKEY_Device_PrinterURL() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB48F35A, &HBE6E, &H4F17, &HB1, &H8, &H3C, &H40, &H73, &HD1, &H66, &H9A, 15)
PKEY_Device_PrinterURL = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_DeviceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 1)
PKEY_DeviceInterface_Bluetooth_DeviceAddress = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_Manufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 4)
PKEY_DeviceInterface_Bluetooth_Manufacturer = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_ModelNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 5)
PKEY_DeviceInterface_Bluetooth_ModelNumber = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_ProductId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 7)
PKEY_DeviceInterface_Bluetooth_ProductId = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_ServiceGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 2)
PKEY_DeviceInterface_Bluetooth_ServiceGuid = pkk
End Function
Public Function PKEY_DeviceInterface_Bluetooth_VendorId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BD67D8B, &H8BEB, &H48D5, &H87, &HE0, &H6C, &HDA, &H34, &H28, &H4, &HA, 6)
PKEY_DeviceInterface_Bluetooth_VendorId = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_IsReadOnly() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 4)
PKEY_DeviceInterface_Hid_IsReadOnly = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_ProductId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 6)
PKEY_DeviceInterface_Hid_ProductId = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_UsageId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 3)
PKEY_DeviceInterface_Hid_UsageId = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_UsagePage() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 2)
PKEY_DeviceInterface_Hid_UsagePage = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_VendorId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 5)
PKEY_DeviceInterface_Hid_VendorId = pkk
End Function
Public Function PKEY_DeviceInterface_Hid_VersionNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCBF38310, &H4A17, &H4310, &HA1, &HEB, &H24, &H7F, &HB, &H67, &H59, &H3B, 7)
PKEY_DeviceInterface_Hid_VersionNumber = pkk
End Function
Public Function PKEY_DeviceInterface_PrinterDriverDirectory() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H847C66DE, &HB8D6, &H4AF9, &HAB, &HC3, &H6F, &H4F, &H92, &H6B, &HC0, &H39, 14)
PKEY_DeviceInterface_PrinterDriverDirectory = pkk
End Function
Public Function PKEY_DeviceInterface_PrinterDriverName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFC47170, &H14F5, &H498C, &H8F, &H30, &HB0, &HD1, &H9B, &HE4, &H49, &HC6, 11)
PKEY_DeviceInterface_PrinterDriverName = pkk
End Function
Public Function PKEY_DeviceInterface_PrinterEnumerationFlag() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA00742A1, &HCD8C, &H4B37, &H95, &HAB, &H70, &H75, &H55, &H87, &H76, &H7A, 3)
PKEY_DeviceInterface_PrinterEnumerationFlag = pkk
End Function
Public Function PKEY_DeviceInterface_PrinterName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA7B84EF, &HC27, &H463F, &H84, &HEF, &H6, &HC5, &H7, &H0, &H1, &HBE, 10)
PKEY_DeviceInterface_PrinterName = pkk
End Function
Public Function PKEY_DeviceInterface_PrinterPortName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEEC7B761, &H6F94, &H41B1, &H94, &H9F, &HC7, &H29, &H72, &HD, &HD1, &H3C, 12)
PKEY_DeviceInterface_PrinterPortName = pkk
End Function
Public Function PKEY_DeviceInterface_Proximity_SupportsNfc() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFB3842CD, &H9E2A, &H4F83, &H8F, &HCC, &H4B, &H7, &H61, &H13, &H9A, &HE9, 2)
PKEY_DeviceInterface_Proximity_SupportsNfc = pkk
End Function
Public Function PKEY_DeviceInterface_Serial_PortName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4C6BF15C, &H4C03, &H4AAC, &H91, &HF5, &H64, &HC0, &HF8, &H52, &HBC, &HF4, 4)
PKEY_DeviceInterface_Serial_PortName = pkk
End Function
Public Function PKEY_DeviceInterface_Serial_UsbProductId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4C6BF15C, &H4C03, &H4AAC, &H91, &HF5, &H64, &HC0, &HF8, &H52, &HBC, &HF4, 3)
PKEY_DeviceInterface_Serial_UsbProductId = pkk
End Function
Public Function PKEY_DeviceInterface_Serial_UsbVendorId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4C6BF15C, &H4C03, &H4AAC, &H91, &HF5, &H64, &HC0, &HF8, &H52, &HBC, &HF4, 2)
PKEY_DeviceInterface_Serial_UsbVendorId = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_DeviceInterfaceClasses() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 7)
PKEY_DeviceInterface_WinUsb_DeviceInterfaceClasses = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_UsbClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 4)
PKEY_DeviceInterface_WinUsb_UsbClass = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_UsbProductId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 3)
PKEY_DeviceInterface_WinUsb_UsbProductId = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_UsbProtocol() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 6)
PKEY_DeviceInterface_WinUsb_UsbProtocol = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_UsbSubClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 5)
PKEY_DeviceInterface_WinUsb_UsbSubClass = pkk
End Function
Public Function PKEY_DeviceInterface_WinUsb_UsbVendorId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95E127B5, &H79CC, &H4E83, &H9C, &H9E, &H84, &H22, &H18, &H7B, &H3E, &HE, 2)
PKEY_DeviceInterface_WinUsb_UsbVendorId = pkk
End Function
Public Function PKEY_Devices_Aep_AepId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3B2CE006, &H5E61, &H4FDE, &HBA, &HB8, &H9B, &H8A, &HAC, &H9B, &H26, &HDF, 8)
PKEY_Devices_Aep_AepId = pkk
End Function
Public Function PKEY_Devices_Aep_CanPair() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE7C3FB29, &HCAA7, &H4F47, &H8C, &H8B, &HBE, &H59, &HB3, &H30, &HD4, &HC5, 3)
PKEY_Devices_Aep_CanPair = pkk
End Function
Public Function PKEY_Devices_Aep_Category() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 17)
PKEY_Devices_Aep_Category = pkk
End Function
Public Function PKEY_Devices_Aep_ContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE7C3FB29, &HCAA7, &H4F47, &H8C, &H8B, &HBE, &H59, &HB3, &H30, &HD4, &HC5, 2)
PKEY_Devices_Aep_ContainerId = pkk
End Function
Public Function PKEY_Devices_Aep_DeviceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 12)
PKEY_Devices_Aep_DeviceAddress = pkk
End Function
Public Function PKEY_Devices_Aep_IsConnected() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 7)
PKEY_Devices_Aep_IsConnected = pkk
End Function
Public Function PKEY_Devices_Aep_IsPaired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 16)
PKEY_Devices_Aep_IsPaired = pkk
End Function
Public Function PKEY_Devices_Aep_IsPresent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 9)
PKEY_Devices_Aep_IsPresent = pkk
End Function
Public Function PKEY_Devices_Aep_Manufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 5)
PKEY_Devices_Aep_Manufacturer = pkk
End Function
Public Function PKEY_Devices_Aep_ModelId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 4)
PKEY_Devices_Aep_ModelId = pkk
End Function
Public Function PKEY_Devices_Aep_ModelName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 3)
PKEY_Devices_Aep_ModelName = pkk
End Function
Public Function PKEY_Devices_Aep_ProtocolId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3B2CE006, &H5E61, &H4FDE, &HBA, &HB8, &H9B, &H8A, &HAC, &H9B, &H26, &HDF, 5)
PKEY_Devices_Aep_ProtocolId = pkk
End Function
Public Function PKEY_Devices_Aep_SignalStrength() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA35996AB, &H11CF, &H4935, &H8B, &H61, &HA6, &H76, &H10, &H81, &HEC, &HDF, 6)
PKEY_Devices_Aep_SignalStrength = pkk
End Function
Public Function PKEY_Devices_AepContainer_CanPair() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 3)
PKEY_Devices_AepContainer_CanPair = pkk
End Function
Public Function PKEY_Devices_AepContainer_Categories() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 9)
PKEY_Devices_AepContainer_Categories = pkk
End Function
Public Function PKEY_Devices_AepContainer_Children() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 2)
PKEY_Devices_AepContainer_Children = pkk
End Function
Public Function PKEY_Devices_AepContainer_ContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 12)
PKEY_Devices_AepContainer_ContainerId = pkk
End Function
Public Function PKEY_Devices_AepContainer_DialProtocol_InstalledApplications() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6AF55D45, &H38DB, &H4495, &HAC, &HB0, &HD4, &H72, &H8A, &H3B, &H83, &H14, 6)
PKEY_Devices_AepContainer_DialProtocol_InstalledApplications = pkk
End Function
Public Function PKEY_Devices_AepContainer_IsPaired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 4)
PKEY_Devices_AepContainer_IsPaired = pkk
End Function
Public Function PKEY_Devices_AepContainer_IsPresent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 11)
PKEY_Devices_AepContainer_IsPresent = pkk
End Function
Public Function PKEY_Devices_AepContainer_Manufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 6)
PKEY_Devices_AepContainer_Manufacturer = pkk
End Function
Public Function PKEY_Devices_AepContainer_ModelIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 8)
PKEY_Devices_AepContainer_ModelIds = pkk
End Function
Public Function PKEY_Devices_AepContainer_ModelName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 7)
PKEY_Devices_AepContainer_ModelName = pkk
End Function
Public Function PKEY_Devices_AepContainer_ProtocolIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBBA1EDE, &H7566, &H4F47, &H90, &HEC, &H25, &HFC, &H56, &H7C, &HED, &H2A, 13)
PKEY_Devices_AepContainer_ProtocolIds = pkk
End Function
Public Function PKEY_Devices_AepContainer_SupportedUriSchemes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6AF55D45, &H38DB, &H4495, &HAC, &HB0, &HD4, &H72, &H8A, &H3B, &H83, &H14, 5)
PKEY_Devices_AepContainer_SupportedUriSchemes = pkk
End Function
Public Function PKEY_Devices_AepContainer_SupportsAudio() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6AF55D45, &H38DB, &H4495, &HAC, &HB0, &HD4, &H72, &H8A, &H3B, &H83, &H14, 2)
PKEY_Devices_AepContainer_SupportsAudio = pkk
End Function
Public Function PKEY_Devices_AepContainer_SupportsImages() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6AF55D45, &H38DB, &H4495, &HAC, &HB0, &HD4, &H72, &H8A, &H3B, &H83, &H14, 4)
PKEY_Devices_AepContainer_SupportsImages = pkk
End Function
Public Function PKEY_Devices_AepContainer_SupportsVideo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6AF55D45, &H38DB, &H4495, &HAC, &HB0, &HD4, &H72, &H8A, &H3B, &H83, &H14, 3)
PKEY_Devices_AepContainer_SupportsVideo = pkk
End Function
Public Function PKEY_Devices_AepService_AepId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9C141A9, &H1B4C, &H4F17, &HA9, &HD1, &HF2, &H98, &H53, &H8C, &HAD, &HB8, 6)
PKEY_Devices_AepService_AepId = pkk
End Function
Public Function PKEY_Devices_AepService_ContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H71724756, &H3E74, &H4432, &H9B, &H59, &HE7, &HB2, &HF6, &H68, &HA5, &H93, 4)
PKEY_Devices_AepService_ContainerId = pkk
End Function
Public Function PKEY_Devices_AepService_FriendlyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H71724756, &H3E74, &H4432, &H9B, &H59, &HE7, &HB2, &HF6, &H68, &HA5, &H93, 2)
PKEY_Devices_AepService_FriendlyName = pkk
End Function
Public Function PKEY_Devices_AepService_ParentAepIsPaired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9C141A9, &H1B4C, &H4F17, &HA9, &HD1, &HF2, &H98, &H53, &H8C, &HAD, &HB8, 7)
PKEY_Devices_AepService_ParentAepIsPaired = pkk
End Function
Public Function PKEY_Devices_AepService_ProtocolId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9C141A9, &H1B4C, &H4F17, &HA9, &HD1, &HF2, &H98, &H53, &H8C, &HAD, &HB8, 5)
PKEY_Devices_AepService_ProtocolId = pkk
End Function
Public Function PKEY_Devices_AepService_ServiceClassId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H71724756, &H3E74, &H4432, &H9B, &H59, &HE7, &HB2, &HF6, &H68, &HA5, &H93, 3)
PKEY_Devices_AepService_ServiceClassId = pkk
End Function
Public Function PKEY_Devices_AepService_ServiceId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9C141A9, &H1B4C, &H4F17, &HA9, &HD1, &HF2, &H98, &H53, &H8C, &HAD, &HB8, 2)
PKEY_Devices_AepService_ServiceId = pkk
End Function
Public Function PKEY_Devices_AppPackageFamilyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H51236583, &HC4A, &H4FE8, &HB8, &H1F, &H16, &H6A, &HEC, &H13, &HF5, &H10, 100)
PKEY_Devices_AppPackageFamilyName = pkk
End Function
Public Function PKEY_Devices_AudioDevice_RawProcessingSupported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8943B373, &H388C, &H4395, &HB5, &H57, &HBC, &H6D, &HBA, &HFF, &HAF, &HDB, 2)
PKEY_Devices_AudioDevice_RawProcessingSupported = pkk
End Function
Public Function PKEY_Devices_AudioDevice_SpeechProcessingSupported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFB1DE864, &HE06D, &H47F4, &H82, &HA6, &H8A, &HA, &HEF, &H44, &H49, &H3C, 2)
PKEY_Devices_AudioDevice_SpeechProcessingSupported = pkk
End Function
Public Function PKEY_Devices_BatteryLife() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 10)
PKEY_Devices_BatteryLife = pkk
End Function
Public Function PKEY_Devices_BatteryPlusCharging() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 22)
PKEY_Devices_BatteryPlusCharging = pkk
End Function
Public Function PKEY_Devices_BatteryPlusChargingText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 23)
PKEY_Devices_BatteryPlusChargingText = pkk
End Function
Public Function PKEY_Devices_Category() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 91)
PKEY_Devices_Category = pkk
End Function
Public Function PKEY_Devices_CategoryGroup() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 94)
PKEY_Devices_CategoryGroup = pkk
End Function
Public Function PKEY_Devices_CategoryIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 90)
PKEY_Devices_CategoryIds = pkk
End Function
Public Function PKEY_Devices_CategoryPlural() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 92)
PKEY_Devices_CategoryPlural = pkk
End Function
Public Function PKEY_Devices_ChargingState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 11)
PKEY_Devices_ChargingState = pkk
End Function
Public Function PKEY_Devices_Children() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 9)
PKEY_Devices_Children = pkk
End Function
Public Function PKEY_Devices_ClassGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 10)
PKEY_Devices_ClassGuid = pkk
End Function
Public Function PKEY_Devices_CompatibleIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 4)
PKEY_Devices_CompatibleIds = pkk
End Function
Public Function PKEY_Devices_Connected() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 55)
PKEY_Devices_Connected = pkk
End Function
Public Function PKEY_Devices_ContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8C7ED206, &H3F8A, &H4827, &HB3, &HAB, &HAE, &H9E, &H1F, &HAE, &HFC, &H6C, 2)
PKEY_Devices_ContainerId = pkk
End Function
Public Function PKEY_Devices_DefaultTooltip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H880F70A2, &H6082, &H47AC, &H8A, &HAB, &HA7, &H39, &HD1, &HA3, &H0, &HC3, 153)
PKEY_Devices_DefaultTooltip = pkk
End Function
Public Function PKEY_Devices_DeviceCapabilities() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 17)
PKEY_Devices_DeviceCapabilities = pkk
End Function
Public Function PKEY_Devices_DeviceCharacteristics() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 29)
PKEY_Devices_DeviceCharacteristics = pkk
End Function
Public Function PKEY_Devices_DeviceDescription1() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 81)
PKEY_Devices_DeviceDescription1 = pkk
End Function
Public Function PKEY_Devices_DeviceDescription2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 82)
PKEY_Devices_DeviceDescription2 = pkk
End Function
Public Function PKEY_Devices_DeviceHasProblem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 6)
PKEY_Devices_DeviceHasProblem = pkk
End Function
Public Function PKEY_Devices_DeviceInstanceId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 256)
PKEY_Devices_DeviceInstanceId = pkk
End Function
Public Function PKEY_Devices_DeviceManufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 13)
PKEY_Devices_DeviceManufacturer = pkk
End Function
Public Function PKEY_Devices_DevObjectType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H13673F42, &HA3D6, &H49F6, &HB4, &HDA, &HAE, &H46, &HE0, &HC5, &H23, &H7C, 2)
PKEY_Devices_DevObjectType = pkk
End Function
Public Function PKEY_Devices_DialProtocol_InstalledApplications() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6845CC72, &H1B71, &H48C3, &HAF, &H86, &HB0, &H91, &H71, &HA1, &H9B, &H14, 3)
PKEY_Devices_DialProtocol_InstalledApplications = pkk
End Function
Public Function PKEY_Devices_DiscoveryMethod() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 52)
PKEY_Devices_DiscoveryMethod = pkk
End Function
Public Function PKEY_Devices_Dnssd_Domain() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 3)
PKEY_Devices_Dnssd_Domain = pkk
End Function
Public Function PKEY_Devices_Dnssd_FullName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 5)
PKEY_Devices_Dnssd_FullName = pkk
End Function
Public Function PKEY_Devices_Dnssd_HostName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 7)
PKEY_Devices_Dnssd_HostName = pkk
End Function
Public Function PKEY_Devices_Dnssd_InstanceName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 4)
PKEY_Devices_Dnssd_InstanceName = pkk
End Function
Public Function PKEY_Devices_Dnssd_NetworkAdapterId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 11)
PKEY_Devices_Dnssd_NetworkAdapterId = pkk
End Function
Public Function PKEY_Devices_Dnssd_PortNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 12)
PKEY_Devices_Dnssd_PortNumber = pkk
End Function
Public Function PKEY_Devices_Dnssd_Priority() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 9)
PKEY_Devices_Dnssd_Priority = pkk
End Function
Public Function PKEY_Devices_Dnssd_ServiceName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 2)
PKEY_Devices_Dnssd_ServiceName = pkk
End Function
Public Function PKEY_Devices_Dnssd_TextAttributes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 6)
PKEY_Devices_Dnssd_TextAttributes = pkk
End Function
Public Function PKEY_Devices_Dnssd_Ttl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 10)
PKEY_Devices_Dnssd_Ttl = pkk
End Function
Public Function PKEY_Devices_Dnssd_Weight() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBF79C0AB, &HBB74, &H4CEE, &HB0, &H70, &H47, &HB, &H5A, &HE2, &H2, &HEA, 8)
PKEY_Devices_Dnssd_Weight = pkk
End Function
Public Function PKEY_Devices_FriendlyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 12288)
PKEY_Devices_FriendlyName = pkk
End Function
Public Function PKEY_Devices_FunctionPaths() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 3)
PKEY_Devices_FunctionPaths = pkk
End Function
Public Function PKEY_Devices_GlyphIcon() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H51236583, &HC4A, &H4FE8, &HB8, &H1F, &H16, &H6A, &HEC, &H13, &HF5, &H10, 123)
PKEY_Devices_GlyphIcon = pkk
End Function
Public Function PKEY_Devices_HardwareIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 3)
PKEY_Devices_HardwareIds = pkk
End Function
Public Function PKEY_Devices_Icon() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 57)
PKEY_Devices_Icon = pkk
End Function
Public Function PKEY_Devices_InLocalMachineContainer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8C7ED206, &H3F8A, &H4827, &HB3, &HAB, &HAE, &H9E, &H1F, &HAE, &HFC, &H6C, 4)
PKEY_Devices_InLocalMachineContainer = pkk
End Function
Public Function PKEY_Devices_InterfaceClassGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 4)
PKEY_Devices_InterfaceClassGuid = pkk
End Function
Public Function PKEY_Devices_InterfaceEnabled() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 3)
PKEY_Devices_InterfaceEnabled = pkk
End Function
Public Function PKEY_Devices_InterfacePaths() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 2)
PKEY_Devices_InterfacePaths = pkk
End Function
Public Function PKEY_Devices_IpAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 12297)
PKEY_Devices_IpAddress = pkk
End Function
Public Function PKEY_Devices_IsDefault() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 86)
PKEY_Devices_IsDefault = pkk
End Function
Public Function PKEY_Devices_IsNetworkConnected() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 85)
PKEY_Devices_IsNetworkConnected = pkk
End Function
Public Function PKEY_Devices_IsShared() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 84)
PKEY_Devices_IsShared = pkk
End Function
Public Function PKEY_Devices_IsSoftwareInstalling() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H83DA6326, &H97A6, &H4088, &H94, &H53, &HA1, &H92, &H3F, &H57, &H3B, &H29, 9)
PKEY_Devices_IsSoftwareInstalling = pkk
End Function
Public Function PKEY_Devices_LaunchDeviceStageFromExplorer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 77)
PKEY_Devices_LaunchDeviceStageFromExplorer = pkk
End Function
Public Function PKEY_Devices_LocalMachine() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 70)
PKEY_Devices_LocalMachine = pkk
End Function
Public Function PKEY_Devices_LocationPaths() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 37)
PKEY_Devices_LocationPaths = pkk
End Function
Public Function PKEY_Devices_Manufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 8192)
PKEY_Devices_Manufacturer = pkk
End Function
Public Function PKEY_Devices_MetadataPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 71)
PKEY_Devices_MetadataPath = pkk
End Function
Public Function PKEY_Devices_MicrophoneArray_Geometry() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA1829EA2, &H27EB, &H459E, &H93, &H5D, &HB2, &HFA, &HD7, &HB0, &H77, &H62, 2)
PKEY_Devices_MicrophoneArray_Geometry = pkk
End Function
Public Function PKEY_Devices_MissedCalls() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 5)
PKEY_Devices_MissedCalls = pkk
End Function
Public Function PKEY_Devices_ModelId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 2)
PKEY_Devices_ModelId = pkk
End Function
Public Function PKEY_Devices_ModelName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 8194)
PKEY_Devices_ModelName = pkk
End Function
Public Function PKEY_Devices_ModelNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 8195)
PKEY_Devices_ModelNumber = pkk
End Function
Public Function PKEY_Devices_NetworkedTooltip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H880F70A2, &H6082, &H47AC, &H8A, &HAB, &HA7, &H39, &HD1, &HA3, &H0, &HC3, 152)
PKEY_Devices_NetworkedTooltip = pkk
End Function
Public Function PKEY_Devices_NetworkName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 7)
PKEY_Devices_NetworkName = pkk
End Function
Public Function PKEY_Devices_NetworkType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 8)
PKEY_Devices_NetworkType = pkk
End Function
Public Function PKEY_Devices_NewPictures() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 4)
PKEY_Devices_NewPictures = pkk
End Function
Public Function PKEY_Devices_Notification() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6704B0C, &HE830, &H4C81, &H91, &H78, &H91, &HE4, &HE9, &H5A, &H80, &HA0, 3)
PKEY_Devices_Notification = pkk
End Function
Public Function PKEY_Devices_Notifications_LowBattery() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC4C07F2B, &H8524, &H4E66, &HAE, &H3A, &HA6, &H23, &H5F, &H10, &H3B, &HEB, 2)
PKEY_Devices_Notifications_LowBattery = pkk
End Function
Public Function PKEY_Devices_Notifications_MissedCall() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6614EF48, &H4EFE, &H4424, &H9E, &HDA, &HC7, &H9F, &H40, &H4E, &HDF, &H3E, 2)
PKEY_Devices_Notifications_MissedCall = pkk
End Function
Public Function PKEY_Devices_Notifications_NewMessage() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BE9260A, &H2012, &H4742, &HA5, &H55, &HF4, &H1B, &H63, &H8B, &H7D, &HCB, 2)
PKEY_Devices_Notifications_NewMessage = pkk
End Function
Public Function PKEY_Devices_Notifications_NewVoicemail() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59569556, &HA08, &H4212, &H95, &HB9, &HFA, &HE2, &HAD, &H64, &H13, &HDB, 2)
PKEY_Devices_Notifications_NewVoicemail = pkk
End Function
Public Function PKEY_Devices_Notifications_StorageFull() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0E00EE1, &HF0C7, &H4D41, &HB8, &HE7, &H26, &HA7, &HBD, &H8D, &H38, &HB0, 2)
PKEY_Devices_Notifications_StorageFull = pkk
End Function
Public Function PKEY_Devices_Notifications_StorageFullLinkText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0E00EE1, &HF0C7, &H4D41, &HB8, &HE7, &H26, &HA7, &HBD, &H8D, &H38, &HB0, 3)
PKEY_Devices_Notifications_StorageFullLinkText = pkk
End Function
Public Function PKEY_Devices_NotificationStore() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6704B0C, &HE830, &H4C81, &H91, &H78, &H91, &HE4, &HE9, &H5A, &H80, &HA0, 2)
PKEY_Devices_NotificationStore = pkk
End Function
Public Function PKEY_Devices_NotWorkingProperly() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 83)
PKEY_Devices_NotWorkingProperly = pkk
End Function
Public Function PKEY_Devices_Paired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 56)
PKEY_Devices_Paired = pkk
End Function
Public Function PKEY_Devices_Parent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 8)
PKEY_Devices_Parent = pkk
End Function
Public Function PKEY_Devices_PhysicalDeviceLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 9)
PKEY_Devices_PhysicalDeviceLocation = pkk
End Function
Public Function PKEY_Devices_PlaybackPositionPercent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3633DE59, &H6825, &H4381, &HA4, &H9B, &H9F, &H6B, &HA1, &H3A, &H14, &H71, 5)
PKEY_Devices_PlaybackPositionPercent = pkk
End Function
Public Function PKEY_Devices_PlaybackState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3633DE59, &H6825, &H4381, &HA4, &H9B, &H9F, &H6B, &HA1, &H3A, &H14, &H71, 2)
PKEY_Devices_PlaybackState = pkk
End Function
Public Function PKEY_Devices_PlaybackTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3633DE59, &H6825, &H4381, &HA4, &H9B, &H9F, &H6B, &HA1, &H3A, &H14, &H71, 3)
PKEY_Devices_PlaybackTitle = pkk
End Function
Public Function PKEY_Devices_Present() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 5)
PKEY_Devices_Present = pkk
End Function
Public Function PKEY_Devices_PresentationUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 8198)
PKEY_Devices_PresentationUrl = pkk
End Function
Public Function PKEY_Devices_PrimaryCategory() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 10)
PKEY_Devices_PrimaryCategory = pkk
End Function
Public Function PKEY_Devices_RemainingDuration() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3633DE59, &H6825, &H4381, &HA4, &H9B, &H9F, &H6B, &HA1, &H3A, &H14, &H71, 4)
PKEY_Devices_RemainingDuration = pkk
End Function
Public Function PKEY_Devices_RestrictedInterface() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 6)
PKEY_Devices_RestrictedInterface = pkk
End Function
Public Function PKEY_Devices_Roaming() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 9)
PKEY_Devices_Roaming = pkk
End Function
Public Function PKEY_Devices_SafeRemovalRequired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFD97640, &H86A3, &H4210, &HB6, &H7C, &H28, &H9C, &H41, &HAA, &HBE, &H55, 2)
PKEY_Devices_SafeRemovalRequired = pkk
End Function
Public Function PKEY_Devices_ServiceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 16384)
PKEY_Devices_ServiceAddress = pkk
End Function
Public Function PKEY_Devices_ServiceId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H656A3BB3, &HECC0, &H43FD, &H84, &H77, &H4A, &HE0, &H40, &H4A, &H96, &HCD, 16385)
PKEY_Devices_ServiceId = pkk
End Function
Public Function PKEY_Devices_SharedTooltip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H880F70A2, &H6082, &H47AC, &H8A, &HAB, &HA7, &H39, &HD1, &HA3, &H0, &HC3, 151)
PKEY_Devices_SharedTooltip = pkk
End Function
Public Function PKEY_Devices_SignalStrength() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 2)
PKEY_Devices_SignalStrength = pkk
End Function
Public Function PKEY_Devices_SmartCards_ReaderKind() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD6B5B883, &H18BD, &H4B4D, &HB2, &HEC, &H9E, &H38, &HAF, &HFE, &HDA, &H82, 2)
PKEY_Devices_SmartCards_ReaderKind = pkk
End Function
Public Function PKEY_Devices_Status() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 259)
PKEY_Devices_Status = pkk
End Function
Public Function PKEY_Devices_Status1() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 257)
PKEY_Devices_Status1 = pkk
End Function
Public Function PKEY_Devices_Status2() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD08DD4C0, &H3A9E, &H462E, &H82, &H90, &H7B, &H63, &H6B, &H25, &H76, &HB9, 258)
PKEY_Devices_Status2 = pkk
End Function
Public Function PKEY_Devices_StorageCapacity() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 12)
PKEY_Devices_StorageCapacity = pkk
End Function
Public Function PKEY_Devices_StorageFreeSpace() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 13)
PKEY_Devices_StorageFreeSpace = pkk
End Function
Public Function PKEY_Devices_StorageFreeSpacePercent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 14)
PKEY_Devices_StorageFreeSpacePercent = pkk
End Function
Public Function PKEY_Devices_TextMessages() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 3)
PKEY_Devices_TextMessages = pkk
End Function
Public Function PKEY_Devices_Voicemail() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49CD1F76, &H5626, &H4B17, &HA4, &HE8, &H18, &HB4, &HAA, &H1A, &H22, &H13, 6)
PKEY_Devices_Voicemail = pkk
End Function
Public Function PKEY_Devices_WiaDeviceType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6BDD1FC6, &H810F, &H11D0, &HBE, &HC7, &H8, &H0, &H2B, &HE2, &H9, &H2F, 2)
PKEY_Devices_WiaDeviceType = pkk
End Function
Public Function PKEY_Devices_WiFi_InterfaceGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEF1167EB, &HCBFC, &H4341, &HA5, &H68, &HA7, &HC9, &H1A, &H68, &H98, &H2C, 2)
PKEY_Devices_WiFi_InterfaceGuid = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_DeviceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 13)
PKEY_Devices_WiFiDirect_DeviceAddress = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_GroupId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 4)
PKEY_Devices_WiFiDirect_GroupId = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_InformationElements() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 12)
PKEY_Devices_WiFiDirect_InformationElements = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_InterfaceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 2)
PKEY_Devices_WiFiDirect_InterfaceAddress = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_InterfaceGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 3)
PKEY_Devices_WiFiDirect_InterfaceGuid = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_IsConnected() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 5)
PKEY_Devices_WiFiDirect_IsConnected = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_IsLegacyDevice() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 7)
PKEY_Devices_WiFiDirect_IsLegacyDevice = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_IsMiracastLcpSupported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 9)
PKEY_Devices_WiFiDirect_IsMiracastLcpSupported = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_IsVisible() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 6)
PKEY_Devices_WiFiDirect_IsVisible = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_MiracastVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 8)
PKEY_Devices_WiFiDirect_MiracastVersion = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_Services() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 10)
PKEY_Devices_WiFiDirect_Services = pkk
End Function
Public Function PKEY_Devices_WiFiDirect_SupportedChannelList() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1506935D, &HE3E7, &H450F, &H86, &H37, &H82, &H23, &H3E, &HBE, &H5F, &H6E, 11)
PKEY_Devices_WiFiDirect_SupportedChannelList = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_AdvertisementId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 5)
PKEY_Devices_WiFiDirectServices_AdvertisementId = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_RequestServiceInformation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 7)
PKEY_Devices_WiFiDirectServices_RequestServiceInformation = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_ServiceAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 2)
PKEY_Devices_WiFiDirectServices_ServiceAddress = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_ServiceConfigMethods() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 6)
PKEY_Devices_WiFiDirectServices_ServiceConfigMethods = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_ServiceInformation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 4)
PKEY_Devices_WiFiDirectServices_ServiceInformation = pkk
End Function
Public Function PKEY_Devices_WiFiDirectServices_ServiceName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H31B37743, &H7C5E, &H4005, &H93, &HE6, &HE9, &H53, &HF9, &H2B, &H82, &HE9, 3)
PKEY_Devices_WiFiDirectServices_ServiceName = pkk
End Function
Public Function PKEY_Devices_WinPhone8CameraFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB7B4D61C, &H5A64, &H4187, &HA5, &H2E, &HB1, &H53, &H9F, &H35, &H90, &H99, 2)
PKEY_Devices_WinPhone8CameraFlags = pkk
End Function
Public Function PKEY_Devices_Wwan_InterfaceGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFF1167EB, &HCBFC, &H4341, &HA5, &H68, &HA7, &HC9, &H1A, &H68, &H98, &H2C, 2)
PKEY_Devices_Wwan_InterfaceGuid = pkk
End Function
Public Function PKEY_Storage_Portable() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4D1EBEE8, &H803, &H4774, &H98, &H42, &HB7, &H7D, &HB5, &H2, &H65, &HE9, 2)
PKEY_Storage_Portable = pkk
End Function
Public Function PKEY_Storage_RemovableMedia() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4D1EBEE8, &H803, &H4774, &H98, &H42, &HB7, &H7D, &HB5, &H2, &H65, &HE9, 3)
PKEY_Storage_RemovableMedia = pkk
End Function
Public Function PKEY_Storage_SystemCritical() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4D1EBEE8, &H803, &H4774, &H98, &H42, &HB7, &H7D, &HB5, &H2, &H65, &HE9, 4)
PKEY_Storage_SystemCritical = pkk
End Function
Public Function PKEY_Document_ByteCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 4)
PKEY_Document_ByteCount = pkk
End Function
Public Function PKEY_Document_CharacterCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 16)
PKEY_Document_CharacterCount = pkk
End Function
Public Function PKEY_Document_ClientID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H276D7BB0, &H5B34, &H4FB0, &HAA, &H4B, &H15, &H8E, &HD1, &H2A, &H18, &H9, 100)
PKEY_Document_ClientID = pkk
End Function
Public Function PKEY_Document_Contributor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF334115E, &HDA1B, &H4509, &H9B, &H3D, &H11, &H95, &H4, &HDC, &H7A, &HBB, 100)
PKEY_Document_Contributor = pkk
End Function
Public Function PKEY_Document_DateCreated() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 12)
PKEY_Document_DateCreated = pkk
End Function
Public Function PKEY_Document_DatePrinted() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 11)
PKEY_Document_DatePrinted = pkk
End Function
Public Function PKEY_Document_DateSaved() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 13)
PKEY_Document_DateSaved = pkk
End Function
Public Function PKEY_Document_Division() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1E005EE6, &HBF27, &H428B, &HB0, &H1C, &H79, &H67, &H6A, &HCD, &H28, &H70, 100)
PKEY_Document_Division = pkk
End Function
Public Function PKEY_Document_DocumentID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE08805C8, &HE395, &H40DF, &H80, &HD2, &H54, &HF0, &HD6, &HC4, &H31, &H54, 100)
PKEY_Document_DocumentID = pkk
End Function
Public Function PKEY_Document_HiddenSlideCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 9)
PKEY_Document_HiddenSlideCount = pkk
End Function
Public Function PKEY_Document_LastAuthor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 8)
PKEY_Document_LastAuthor = pkk
End Function
Public Function PKEY_Document_LineCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 5)
PKEY_Document_LineCount = pkk
End Function
Public Function PKEY_Document_Manager() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 14)
PKEY_Document_Manager = pkk
End Function
Public Function PKEY_Document_MultimediaClipCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 10)
PKEY_Document_MultimediaClipCount = pkk
End Function
Public Function PKEY_Document_NoteCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 8)
PKEY_Document_NoteCount = pkk
End Function
Public Function PKEY_Document_PageCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 14)
PKEY_Document_PageCount = pkk
End Function
Public Function PKEY_Document_ParagraphCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 6)
PKEY_Document_ParagraphCount = pkk
End Function
Public Function PKEY_Document_PresentationFormat() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 3)
PKEY_Document_PresentationFormat = pkk
End Function
Public Function PKEY_Document_RevisionNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 9)
PKEY_Document_RevisionNumber = pkk
End Function
Public Function PKEY_Document_Security() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 19)
PKEY_Document_Security = pkk
End Function
Public Function PKEY_Document_SlideCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 7)
PKEY_Document_SlideCount = pkk
End Function
Public Function PKEY_Document_Template() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 7)
PKEY_Document_Template = pkk
End Function
Public Function PKEY_Document_TotalEditingTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 10)
PKEY_Document_TotalEditingTime = pkk
End Function
Public Function PKEY_Document_Version() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD5CDD502, &H2E9C, &H101B, &H93, &H97, &H8, &H0, &H2B, &H2C, &HF9, &HAE, 29)
PKEY_Document_Version = pkk
End Function
Public Function PKEY_Document_WordCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF29F85E0, &H4FF9, &H1068, &HAB, &H91, &H8, &H0, &H2B, &H27, &HB3, &HD9, 15)
PKEY_Document_WordCount = pkk
End Function
Public Function PKEY_DRM_DatePlayExpires() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 6)
PKEY_DRM_DatePlayExpires = pkk
End Function
Public Function PKEY_DRM_DatePlayStarts() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 5)
PKEY_DRM_DatePlayStarts = pkk
End Function
Public Function PKEY_DRM_Description() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 3)
PKEY_DRM_Description = pkk
End Function
Public Function PKEY_DRM_IsDisabled() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 7)
PKEY_DRM_IsDisabled = pkk
End Function
Public Function PKEY_DRM_IsProtected() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 2)
PKEY_DRM_IsProtected = pkk
End Function
Public Function PKEY_DRM_PlayCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAEAC19E4, &H89AE, &H4508, &HB9, &HB7, &HBB, &H86, &H7A, &HBE, &HE2, &HED, 4)
PKEY_DRM_PlayCount = pkk
End Function
Public Function PKEY_GPS_Altitude() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H827EDB4F, &H5B73, &H44A7, &H89, &H1D, &HFD, &HFF, &HAB, &HEA, &H35, &HCA, 100)
PKEY_GPS_Altitude = pkk
End Function
Public Function PKEY_GPS_AltitudeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78342DCB, &HE358, &H4145, &HAE, &H9A, &H6B, &HFE, &H4E, &HF, &H9F, &H51, 100)
PKEY_GPS_AltitudeDenominator = pkk
End Function
Public Function PKEY_GPS_AltitudeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2DAD1EB7, &H816D, &H40D3, &H9E, &HC3, &HC9, &H77, &H3B, &HE2, &HAA, &HDE, 100)
PKEY_GPS_AltitudeNumerator = pkk
End Function
Public Function PKEY_GPS_AltitudeRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H46AC629D, &H75EA, &H4515, &H86, &H7F, &H6D, &HC4, &H32, &H1C, &H58, &H44, 100)
PKEY_GPS_AltitudeRef = pkk
End Function
Public Function PKEY_GPS_AreaInformation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H972E333E, &HAC7E, &H49F1, &H8A, &HDF, &HA7, &HD, &H7, &HA9, &HBC, &HAB, 100)
PKEY_GPS_AreaInformation = pkk
End Function
Public Function PKEY_GPS_Date() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3602C812, &HF3B, &H45F0, &H85, &HAD, &H60, &H34, &H68, &HD6, &H94, &H23, 100)
PKEY_GPS_Date = pkk
End Function
Public Function PKEY_GPS_DestBearing() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC66D4B3C, &HE888, &H47CC, &HB9, &H9F, &H9D, &HCA, &H3E, &HE3, &H4D, &HEA, 100)
PKEY_GPS_DestBearing = pkk
End Function
Public Function PKEY_GPS_DestBearingDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7ABCF4F8, &H7C3F, &H4988, &HAC, &H91, &H8D, &H2C, &H2E, &H97, &HEC, &HA5, 100)
PKEY_GPS_DestBearingDenominator = pkk
End Function
Public Function PKEY_GPS_DestBearingNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBA3B1DA9, &H86EE, &H4B5D, &HA2, &HA4, &HA2, &H71, &HA4, &H29, &HF0, &HCF, 100)
PKEY_GPS_DestBearingNumerator = pkk
End Function
Public Function PKEY_GPS_DestBearingRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9AB84393, &H2A0F, &H4B75, &HBB, &H22, &H72, &H79, &H78, &H69, &H77, &HCB, 100)
PKEY_GPS_DestBearingRef = pkk
End Function
Public Function PKEY_GPS_DestDistance() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA93EAE04, &H6804, &H4F24, &HAC, &H81, &H9, &HB2, &H66, &H45, &H21, &H18, 100)
PKEY_GPS_DestDistance = pkk
End Function
Public Function PKEY_GPS_DestDistanceDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9BC2C99B, &HAC71, &H4127, &H9D, &H1C, &H25, &H96, &HD0, &HD7, &HDC, &HB7, 100)
PKEY_GPS_DestDistanceDenominator = pkk
End Function
Public Function PKEY_GPS_DestDistanceNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2BDA47DA, &H8C6, &H4FE1, &H80, &HBC, &HA7, &H2F, &HC5, &H17, &HC5, &HD0, 100)
PKEY_GPS_DestDistanceNumerator = pkk
End Function
Public Function PKEY_GPS_DestDistanceRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HED4DF2D3, &H8695, &H450B, &H85, &H6F, &HF5, &HC1, &HC5, &H3A, &HCB, &H66, 100)
PKEY_GPS_DestDistanceRef = pkk
End Function
Public Function PKEY_GPS_DestLatitude() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9D1D7CC5, &H5C39, &H451C, &H86, &HB3, &H92, &H8E, &H2D, &H18, &HCC, &H47, 100)
PKEY_GPS_DestLatitude = pkk
End Function
Public Function PKEY_GPS_DestLatitudeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3A372292, &H7FCA, &H49A7, &H99, &HD5, &HE4, &H7B, &HB2, &HD4, &HE7, &HAB, 100)
PKEY_GPS_DestLatitudeDenominator = pkk
End Function
Public Function PKEY_GPS_DestLatitudeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HECF4B6F6, &HD5A6, &H433C, &HBB, &H92, &H40, &H76, &H65, &HF, &HC8, &H90, 100)
PKEY_GPS_DestLatitudeNumerator = pkk
End Function
Public Function PKEY_GPS_DestLatitudeRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEA820B9, &HCE61, &H4885, &HA1, &H28, &H0, &H5D, &H90, &H87, &HC1, &H92, 100)
PKEY_GPS_DestLatitudeRef = pkk
End Function
Public Function PKEY_GPS_DestLongitude() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H47A96261, &HCB4C, &H4807, &H8A, &HD3, &H40, &HB9, &HD9, &HDB, &HC6, &HBC, 100)
PKEY_GPS_DestLongitude = pkk
End Function
Public Function PKEY_GPS_DestLongitudeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H425D69E5, &H48AD, &H4900, &H8D, &H80, &H6E, &HB6, &HB8, &HD0, &HAC, &H86, 100)
PKEY_GPS_DestLongitudeDenominator = pkk
End Function
Public Function PKEY_GPS_DestLongitudeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA3250282, &HFB6D, &H48D5, &H9A, &H89, &HDB, &HCA, &HCE, &H75, &HCC, &HCF, 100)
PKEY_GPS_DestLongitudeNumerator = pkk
End Function
Public Function PKEY_GPS_DestLongitudeRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H182C1EA6, &H7C1C, &H4083, &HAB, &H4B, &HAC, &H6C, &H9F, &H4E, &HD1, &H28, 100)
PKEY_GPS_DestLongitudeRef = pkk
End Function
Public Function PKEY_GPS_Differential() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAAF4EE25, &HBD3B, &H4DD7, &HBF, &HC4, &H47, &HF7, &H7B, &HB0, &HF, &H6D, 100)
PKEY_GPS_Differential = pkk
End Function
Public Function PKEY_GPS_DOP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF8FB02, &H1837, &H42F1, &HA6, &H97, &HA7, &H1, &H7A, &HA2, &H89, &HB9, 100)
PKEY_GPS_DOP = pkk
End Function
Public Function PKEY_GPS_DOPDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0BE94C5, &H50BA, &H487B, &HBD, &H35, &H6, &H54, &HBE, &H88, &H81, &HED, 100)
PKEY_GPS_DOPDenominator = pkk
End Function
Public Function PKEY_GPS_DOPNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H47166B16, &H364F, &H4AA0, &H9F, &H31, &HE2, &HAB, &H3D, &HF4, &H49, &HC3, 100)
PKEY_GPS_DOPNumerator = pkk
End Function
Public Function PKEY_GPS_ImgDirection() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H16473C91, &HD017, &H4ED9, &HBA, &H4D, &HB6, &HBA, &HA5, &H5D, &HBC, &HF8, 100)
PKEY_GPS_ImgDirection = pkk
End Function
Public Function PKEY_GPS_ImgDirectionDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10B24595, &H41A2, &H4E20, &H93, &HC2, &H57, &H61, &HC1, &H39, &H5F, &H32, 100)
PKEY_GPS_ImgDirectionDenominator = pkk
End Function
Public Function PKEY_GPS_ImgDirectionNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDC5877C7, &H225F, &H45F7, &HBA, &HC7, &HE8, &H13, &H34, &HB6, &H13, &HA, 100)
PKEY_GPS_ImgDirectionNumerator = pkk
End Function
Public Function PKEY_GPS_ImgDirectionRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA4AAA5B7, &H1AD0, &H445F, &H81, &H1A, &HF, &H8F, &H6E, &H67, &HF6, &HB5, 100)
PKEY_GPS_ImgDirectionRef = pkk
End Function
Public Function PKEY_GPS_Latitude() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8727CFFF, &H4868, &H4EC6, &HAD, &H5B, &H81, &HB9, &H85, &H21, &HD1, &HAB, 100)
PKEY_GPS_Latitude = pkk
End Function
Public Function PKEY_GPS_LatitudeDecimal() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF55CDE2, &H4F49, &H450D, &H92, &HC1, &HDC, &HD1, &H63, &H1, &HB1, &HB7, 100)
PKEY_GPS_LatitudeDecimal = pkk
End Function
Public Function PKEY_GPS_LatitudeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H16E634EE, &H2BFF, &H497B, &HBD, &H8A, &H43, &H41, &HAD, &H39, &HEE, &HB9, 100)
PKEY_GPS_LatitudeDenominator = pkk
End Function
Public Function PKEY_GPS_LatitudeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7DDAAAD1, &HCCC8, &H41AE, &HB7, &H50, &HB2, &HCB, &H80, &H31, &HAE, &HA2, 100)
PKEY_GPS_LatitudeNumerator = pkk
End Function
Public Function PKEY_GPS_LatitudeRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H29C0252, &H5B86, &H46C7, &HAC, &HA0, &H27, &H69, &HFF, &HC8, &HE3, &HD4, 100)
PKEY_GPS_LatitudeRef = pkk
End Function
Public Function PKEY_GPS_Longitude() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC4C4DBB2, &HB593, &H466B, &HBB, &HDA, &HD0, &H3D, &H27, &HD5, &HE4, &H3A, 100)
PKEY_GPS_Longitude = pkk
End Function
Public Function PKEY_GPS_LongitudeDecimal() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4679C1B5, &H844D, &H4590, &HBA, &HF5, &HF3, &H22, &H23, &H1F, &H1B, &H81, 100)
PKEY_GPS_LongitudeDecimal = pkk
End Function
Public Function PKEY_GPS_LongitudeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBE6E176C, &H4534, &H4D2C, &HAC, &HE5, &H31, &HDE, &HDA, &HC1, &H60, &H6B, 100)
PKEY_GPS_LongitudeDenominator = pkk
End Function
Public Function PKEY_GPS_LongitudeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2B0F689, &HA914, &H4E45, &H82, &H1D, &H1D, &HDA, &H45, &H2E, &HD2, &HC4, 100)
PKEY_GPS_LongitudeNumerator = pkk
End Function
Public Function PKEY_GPS_LongitudeRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H33DCF22B, &H28D5, &H464C, &H80, &H35, &H1E, &HE9, &HEF, &HD2, &H52, &H78, 100)
PKEY_GPS_LongitudeRef = pkk
End Function
Public Function PKEY_GPS_MapDatum() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2CA2DAE6, &HEDDC, &H407D, &HBE, &HF1, &H77, &H39, &H42, &HAB, &HFA, &H95, 100)
PKEY_GPS_MapDatum = pkk
End Function
Public Function PKEY_GPS_MeasureMode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA015ED5D, &HAAEA, &H4D58, &H8A, &H86, &H3C, &H58, &H69, &H20, &HEA, &HB, 100)
PKEY_GPS_MeasureMode = pkk
End Function
Public Function PKEY_GPS_ProcessingMethod() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59D49E61, &H840F, &H4AA9, &HA9, &H39, &HE2, &H9, &H9B, &H7F, &H63, &H99, 100)
PKEY_GPS_ProcessingMethod = pkk
End Function
Public Function PKEY_GPS_Satellites() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H467EE575, &H1F25, &H4557, &HAD, &H4E, &HB8, &HB5, &H8B, &HD, &H9C, &H15, 100)
PKEY_GPS_Satellites = pkk
End Function
Public Function PKEY_GPS_Speed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDA5D0862, &H6E76, &H4E1B, &HBA, &HBD, &H70, &H2, &H1B, &HD2, &H54, &H94, 100)
PKEY_GPS_Speed = pkk
End Function
Public Function PKEY_GPS_SpeedDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7D122D5A, &HAE5E, &H4335, &H88, &H41, &HD7, &H1E, &H7C, &HE7, &H2F, &H53, 100)
PKEY_GPS_SpeedDenominator = pkk
End Function
Public Function PKEY_GPS_SpeedNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HACC9CE3D, &HC213, &H4942, &H8B, &H48, &H6D, &H8, &H20, &HF2, &H1C, &H6D, 100)
PKEY_GPS_SpeedNumerator = pkk
End Function
Public Function PKEY_GPS_SpeedRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HECF7F4C9, &H544F, &H4D6D, &H9D, &H98, &H8A, &HD7, &H9A, &HDA, &HF4, &H53, 100)
PKEY_GPS_SpeedRef = pkk
End Function
Public Function PKEY_GPS_Status() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H125491F4, &H818F, &H46B2, &H91, &HB5, &HD5, &H37, &H75, &H36, &H17, &HB2, 100)
PKEY_GPS_Status = pkk
End Function
Public Function PKEY_GPS_Track() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H76C09943, &H7C33, &H49E3, &H9E, &H7E, &HCD, &HBA, &H87, &H2C, &HFA, &HDA, 100)
PKEY_GPS_Track = pkk
End Function
Public Function PKEY_GPS_TrackDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC8D1920C, &H1F6, &H40C0, &HAC, &H86, &H2F, &H3A, &H4A, &HD0, &H7, &H70, 100)
PKEY_GPS_TrackDenominator = pkk
End Function
Public Function PKEY_GPS_TrackNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H702926F4, &H44A6, &H43E1, &HAE, &H71, &H45, &H62, &H71, &H16, &H89, &H3B, 100)
PKEY_GPS_TrackNumerator = pkk
End Function
Public Function PKEY_GPS_TrackRef() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H35DBE6FE, &H44C3, &H4400, &HAA, &HAE, &HD2, &HC7, &H99, &HC4, &H7, &HE8, 100)
PKEY_GPS_TrackRef = pkk
End Function
Public Function PKEY_GPS_VersionID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H22704DA4, &HC6B2, &H4A99, &H8E, &H56, &HF1, &H6D, &HF8, &HC9, &H25, &H99, 100)
PKEY_GPS_VersionID = pkk
End Function
Public Function PKEY_History_VisitCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5CBF2787, &H48CF, &H4208, &HB9, &HE, &HEE, &H5E, &H5D, &H42, &H2, &H94, 7)
PKEY_History_VisitCount = pkk
End Function
Public Function PKEY_Image_BitDepth() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 7)
PKEY_Image_BitDepth = pkk
End Function
Public Function PKEY_Image_ColorSpace() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 40961)
PKEY_Image_ColorSpace = pkk
End Function
Public Function PKEY_Image_CompressedBitsPerPixel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H364B6FA9, &H37AB, &H482A, &HBE, &H2B, &HAE, &H2, &HF6, &HD, &H43, &H18, 100)
PKEY_Image_CompressedBitsPerPixel = pkk
End Function
Public Function PKEY_Image_CompressedBitsPerPixelDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1F8844E1, &H24AD, &H4508, &H9D, &HFD, &H53, &H26, &HA4, &H15, &HCE, &H2, 100)
PKEY_Image_CompressedBitsPerPixelDenominator = pkk
End Function
Public Function PKEY_Image_CompressedBitsPerPixelNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD21A7148, &HD32C, &H4624, &H89, &H0, &H27, &H72, &H10, &HF7, &H9C, &HF, 100)
PKEY_Image_CompressedBitsPerPixelNumerator = pkk
End Function
Public Function PKEY_Image_Compression() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 259)
PKEY_Image_Compression = pkk
End Function
Public Function PKEY_Image_CompressionText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3F08E66F, &H2F44, &H4BB9, &HA6, &H82, &HAC, &H35, &HD2, &H56, &H23, &H22, 100)
PKEY_Image_CompressionText = pkk
End Function
Public Function PKEY_Image_Dimensions() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 13)
PKEY_Image_Dimensions = pkk
End Function
Public Function PKEY_Image_HorizontalResolution() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 5)
PKEY_Image_HorizontalResolution = pkk
End Function
Public Function PKEY_Image_HorizontalSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 3)
PKEY_Image_HorizontalSize = pkk
End Function
Public Function PKEY_Image_ImageID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H10DABE05, &H32AA, &H4C29, &HBF, &H1A, &H63, &HE2, &HD2, &H20, &H58, &H7F, 100)
PKEY_Image_ImageID = pkk
End Function
Public Function PKEY_Image_ResolutionUnit() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H19B51FA6, &H1F92, &H4A5C, &HAB, &H48, &H7D, &HF0, &HAB, &HD6, &H74, &H44, 100)
PKEY_Image_ResolutionUnit = pkk
End Function
Public Function PKEY_Image_VerticalResolution() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 6)
PKEY_Image_VerticalResolution = pkk
End Function
Public Function PKEY_Image_VerticalSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 4)
PKEY_Image_VerticalSize = pkk
End Function
Public Function PKEY_Journal_Contacts() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDEA7C82C, &H1D89, &H4A66, &H94, &H27, &HA4, &HE3, &HDE, &HBA, &HBC, &HB1, 100)
PKEY_Journal_Contacts = pkk
End Function
Public Function PKEY_Journal_EntryType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H95BEB1FC, &H326D, &H4644, &HB3, &H96, &HCD, &H3E, &HD9, &HE, &H6D, &HDF, 100)
PKEY_Journal_EntryType = pkk
End Function
Public Function PKEY_LayoutPattern_ContentViewModeForBrowse() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 500)
PKEY_LayoutPattern_ContentViewModeForBrowse = pkk
End Function
Public Function PKEY_LayoutPattern_ContentViewModeForSearch() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 501)
PKEY_LayoutPattern_ContentViewModeForSearch = pkk
End Function
Public Function PKEY_History_SelectionCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1CE0D6BC, &H536C, &H4600, &HB0, &HDD, &H7E, &HC, &H66, &HB3, &H50, &HD5, 8)
PKEY_History_SelectionCount = pkk
End Function
Public Function PKEY_History_TargetUrlHostName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1CE0D6BC, &H536C, &H4600, &HB0, &HDD, &H7E, &HC, &H66, &HB3, &H50, &HD5, 9)
PKEY_History_TargetUrlHostName = pkk
End Function
Public Function PKEY_Link_Arguments() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H436F2667, &H14E2, &H4FEB, &HB3, &HA, &H14, &H6C, &H53, &HB5, &HB6, &H74, 100)
PKEY_Link_Arguments = pkk
End Function
Public Function PKEY_Link_Comment() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB9B4B3FC, &H2B51, &H4A42, &HB5, &HD8, &H32, &H41, &H46, &HAF, &HCF, &H25, 5)
PKEY_Link_Comment = pkk
End Function
Public Function PKEY_Link_DateVisited() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5CBF2787, &H48CF, &H4208, &HB9, &HE, &HEE, &H5E, &H5D, &H42, &H2, &H94, 23)
PKEY_Link_DateVisited = pkk
End Function
Public Function PKEY_Link_Description() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5CBF2787, &H48CF, &H4208, &HB9, &HE, &HEE, &H5E, &H5D, &H42, &H2, &H94, 21)
PKEY_Link_Description = pkk
End Function
Public Function PKEY_Link_FeedItemLocalId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8A2F99F9, &H3C37, &H465D, &HA8, &HD7, &H69, &H77, &H7A, &H24, &H6D, &HC, 2)
PKEY_Link_FeedItemLocalId = pkk
End Function
Public Function PKEY_Link_Status() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB9B4B3FC, &H2B51, &H4A42, &HB5, &HD8, &H32, &H41, &H46, &HAF, &HCF, &H25, 3)
PKEY_Link_Status = pkk
End Function
Public Function PKEY_Link_TargetExtension() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7A7D76F4, &HB630, &H4BD7, &H95, &HFF, &H37, &HCC, &H51, &HA9, &H75, &HC9, 2)
PKEY_Link_TargetExtension = pkk
End Function
Public Function PKEY_Link_TargetParsingPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB9B4B3FC, &H2B51, &H4A42, &HB5, &HD8, &H32, &H41, &H46, &HAF, &HCF, &H25, 2)
PKEY_Link_TargetParsingPath = pkk
End Function
Public Function PKEY_Link_TargetSFGAOFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB9B4B3FC, &H2B51, &H4A42, &HB5, &HD8, &H32, &H41, &H46, &HAF, &HCF, &H25, 8)
PKEY_Link_TargetSFGAOFlags = pkk
End Function
Public Function PKEY_Link_TargetUrlHostName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8A2F99F9, &H3C37, &H465D, &HA8, &HD7, &H69, &H77, &H7A, &H24, &H6D, &HC, 5)
PKEY_Link_TargetUrlHostName = pkk
End Function
Public Function PKEY_Link_TargetUrlPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8A2F99F9, &H3C37, &H465D, &HA8, &HD7, &H69, &H77, &H7A, &H24, &H6D, &HC, 6)
PKEY_Link_TargetUrlPath = pkk
End Function
Public Function PKEY_Media_AuthorUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 32)
PKEY_Media_AuthorUrl = pkk
End Function
Public Function PKEY_Media_AverageLevel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9EDD5B6, &HB301, &H43C5, &H99, &H90, &HD0, &H3, &H2, &HEF, &HFD, &H46, 100)
PKEY_Media_AverageLevel = pkk
End Function
Public Function PKEY_Media_ClassPrimaryID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 13)
PKEY_Media_ClassPrimaryID = pkk
End Function
Public Function PKEY_Media_ClassSecondaryID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 14)
PKEY_Media_ClassSecondaryID = pkk
End Function
Public Function PKEY_Media_CollectionGroupID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 24)
PKEY_Media_CollectionGroupID = pkk
End Function
Public Function PKEY_Media_CollectionID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 25)
PKEY_Media_CollectionID = pkk
End Function
Public Function PKEY_Media_ContentDistributor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 18)
PKEY_Media_ContentDistributor = pkk
End Function
Public Function PKEY_Media_ContentID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 26)
PKEY_Media_ContentID = pkk
End Function
Public Function PKEY_Media_CreatorApplication() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 27)
PKEY_Media_CreatorApplication = pkk
End Function
Public Function PKEY_Media_CreatorApplicationVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 28)
PKEY_Media_CreatorApplicationVersion = pkk
End Function
Public Function PKEY_Media_DateEncoded() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2E4B640D, &H5019, &H46D8, &H88, &H81, &H55, &H41, &H4C, &HC5, &HCA, &HA0, 100)
PKEY_Media_DateEncoded = pkk
End Function
Public Function PKEY_Media_DateReleased() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDE41CC29, &H6971, &H4290, &HB4, &H72, &HF5, &H9F, &H2E, &H2F, &H31, &HE2, 100)
PKEY_Media_DateReleased = pkk
End Function
Public Function PKEY_Media_DlnaProfileID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCFA31B45, &H525D, &H4998, &HBB, &H44, &H3F, &H7D, &H81, &H54, &H2F, &HA4, 100)
PKEY_Media_DlnaProfileID = pkk
End Function
Public Function PKEY_Media_Duration() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440490, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 3)
PKEY_Media_Duration = pkk
End Function
Public Function PKEY_Media_DVDID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 15)
PKEY_Media_DVDID = pkk
End Function
Public Function PKEY_Media_EncodedBy() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 36)
PKEY_Media_EncodedBy = pkk
End Function
Public Function PKEY_Media_EncodingSettings() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 37)
PKEY_Media_EncodingSettings = pkk
End Function
Public Function PKEY_Media_EpisodeNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 100)
PKEY_Media_EpisodeNumber = pkk
End Function
Public Function PKEY_Media_FrameCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6444048F, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 12)
PKEY_Media_FrameCount = pkk
End Function
Public Function PKEY_Media_MCDI() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 16)
PKEY_Media_MCDI = pkk
End Function
Public Function PKEY_Media_MetadataContentProvider() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 17)
PKEY_Media_MetadataContentProvider = pkk
End Function
Public Function PKEY_Media_Producer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 22)
PKEY_Media_Producer = pkk
End Function
Public Function PKEY_Media_PromotionUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 33)
PKEY_Media_PromotionUrl = pkk
End Function
Public Function PKEY_Media_ProtectionType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 38)
PKEY_Media_ProtectionType = pkk
End Function
Public Function PKEY_Media_ProviderRating() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 39)
PKEY_Media_ProviderRating = pkk
End Function
Public Function PKEY_Media_ProviderStyle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 40)
PKEY_Media_ProviderStyle = pkk
End Function
Public Function PKEY_Media_Publisher() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 30)
PKEY_Media_Publisher = pkk
End Function
Public Function PKEY_Media_SeasonNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 101)
PKEY_Media_SeasonNumber = pkk
End Function
Public Function PKEY_Media_SeriesName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 42)
PKEY_Media_SeriesName = pkk
End Function
Public Function PKEY_Media_SubscriptionContentId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9AEBAE7A, &H9644, &H487D, &HA9, &H2C, &H65, &H75, &H85, &HED, &H75, &H1A, 100)
PKEY_Media_SubscriptionContentId = pkk
End Function
Public Function PKEY_Media_SubTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 38)
PKEY_Media_SubTitle = pkk
End Function
Public Function PKEY_Media_ThumbnailLargePath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 47)
PKEY_Media_ThumbnailLargePath = pkk
End Function
Public Function PKEY_Media_ThumbnailLargeUri() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 48)
PKEY_Media_ThumbnailLargeUri = pkk
End Function
Public Function PKEY_Media_ThumbnailSmallPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 49)
PKEY_Media_ThumbnailSmallPath = pkk
End Function
Public Function PKEY_Media_ThumbnailSmallUri() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 50)
PKEY_Media_ThumbnailSmallUri = pkk
End Function
Public Function PKEY_Media_UniqueFileIdentifier() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 35)
PKEY_Media_UniqueFileIdentifier = pkk
End Function
Public Function PKEY_Media_UserNoAutoInfo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 41)
PKEY_Media_UserNoAutoInfo = pkk
End Function
Public Function PKEY_Media_UserWebUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 34)
PKEY_Media_UserWebUrl = pkk
End Function
Public Function PKEY_Media_Writer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 23)
PKEY_Media_Writer = pkk
End Function
Public Function PKEY_Media_Year() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 5)
PKEY_Media_Year = pkk
End Function
Public Function PKEY_Message_AttachmentContents() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H3143BF7C, &H80A8, &H4854, &H88, &H80, &HE2, &HE4, &H1, &H89, &HBD, &HD0, 100)
PKEY_Message_AttachmentContents = pkk
End Function
Public Function PKEY_Message_AttachmentNames() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 21)
PKEY_Message_AttachmentNames = pkk
End Function
Public Function PKEY_Message_BccAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 2)
PKEY_Message_BccAddress = pkk
End Function
Public Function PKEY_Message_BccName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 3)
PKEY_Message_BccName = pkk
End Function
Public Function PKEY_Message_CcAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 4)
PKEY_Message_CcAddress = pkk
End Function
Public Function PKEY_Message_CcName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 5)
PKEY_Message_CcName = pkk
End Function
Public Function PKEY_Message_ConversationID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDC8F80BD, &HAF1E, &H4289, &H85, &HB6, &H3D, &HFC, &H1B, &H49, &H39, &H92, 100)
PKEY_Message_ConversationID = pkk
End Function
Public Function PKEY_Message_ConversationIndex() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDC8F80BD, &HAF1E, &H4289, &H85, &HB6, &H3D, &HFC, &H1B, &H49, &H39, &H92, 101)
PKEY_Message_ConversationIndex = pkk
End Function
Public Function PKEY_Message_DateReceived() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 20)
PKEY_Message_DateReceived = pkk
End Function
Public Function PKEY_Message_DateSent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 19)
PKEY_Message_DateSent = pkk
End Function
Public Function PKEY_Message_Flags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA82D9EE7, &HCA67, &H4312, &H96, &H5E, &H22, &H6B, &HCE, &HA8, &H50, &H23, 100)
PKEY_Message_Flags = pkk
End Function
Public Function PKEY_Message_FromAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 13)
PKEY_Message_FromAddress = pkk
End Function
Public Function PKEY_Message_FromName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 14)
PKEY_Message_FromName = pkk
End Function
Public Function PKEY_Message_HasAttachments() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9C1FCF74, &H2D97, &H41BA, &HB4, &HAE, &HCB, &H2E, &H36, &H61, &HA6, &HE4, 8)
PKEY_Message_HasAttachments = pkk
End Function
Public Function PKEY_Message_IsFwdOrReply() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9A9BC088, &H4F6D, &H469E, &H99, &H19, &HE7, &H5, &H41, &H20, &H40, &HF9, 100)
PKEY_Message_IsFwdOrReply = pkk
End Function
Public Function PKEY_Message_MessageClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCD9ED458, &H8CE, &H418F, &HA7, &HE, &HF9, &H12, &HC7, &HBB, &H9C, &H5C, 103)
PKEY_Message_MessageClass = pkk
End Function
Public Function PKEY_Message_Participants() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A9BA605, &H8E7C, &H4D11, &HAD, &H7D, &HA5, &HA, &HDA, &H18, &HBA, &H1B, 2)
PKEY_Message_Participants = pkk
End Function
Public Function PKEY_Message_ProofInProgress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9098F33C, &H9A7D, &H48A8, &H8D, &HE5, &H2E, &H12, &H27, &HA6, &H4E, &H91, 100)
PKEY_Message_ProofInProgress = pkk
End Function
Public Function PKEY_Message_SenderAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBE1C8E7, &H1981, &H4676, &HAE, &H14, &HFD, &HD7, &H8F, &H5, &HA6, &HE7, 100)
PKEY_Message_SenderAddress = pkk
End Function
Public Function PKEY_Message_SenderName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDA41CFA, &HD224, &H4A18, &HAE, &H2F, &H59, &H61, &H58, &HDB, &H4B, &H3A, 100)
PKEY_Message_SenderName = pkk
End Function
Public Function PKEY_Message_Store() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 15)
PKEY_Message_Store = pkk
End Function
Public Function PKEY_Message_ToAddress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 16)
PKEY_Message_ToAddress = pkk
End Function
Public Function PKEY_Message_ToDoFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1F856A9F, &H6900, &H4ABA, &H95, &H5, &H2D, &H5F, &H1B, &H4D, &H66, &HCB, 100)
PKEY_Message_ToDoFlags = pkk
End Function
Public Function PKEY_Message_ToDoTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBCCC8A3C, &H8CEF, &H42E5, &H9B, &H1C, &HC6, &H90, &H79, &H39, &H8B, &HC7, 100)
PKEY_Message_ToDoTitle = pkk
End Function
Public Function PKEY_Message_ToName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3E0584C, &HB788, &H4A5A, &HBB, &H20, &H7F, &H5A, &H44, &HC9, &HAC, &HDD, 17)
PKEY_Message_ToName = pkk
End Function
Public Function PKEY_Music_AlbumArtist() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 13)
PKEY_Music_AlbumArtist = pkk
End Function
Public Function PKEY_Music_AlbumArtistSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF1FDB4AF, &HF78C, &H466C, &HBB, &H5, &H56, &HE9, &H2D, &HB0, &HB8, &HEC, 103)
PKEY_Music_AlbumArtistSortOverride = pkk
End Function
Public Function PKEY_Music_AlbumID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 100)
PKEY_Music_AlbumID = pkk
End Function
Public Function PKEY_Music_AlbumTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 4)
PKEY_Music_AlbumTitle = pkk
End Function
Public Function PKEY_Music_AlbumTitleSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H13EB7FFC, &HEC89, &H4346, &HB1, &H9D, &HCC, &HC6, &HF1, &H78, &H42, &H23, 101)
PKEY_Music_AlbumTitleSortOverride = pkk
End Function
Public Function PKEY_Music_Artist() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 2)
PKEY_Music_Artist = pkk
End Function
Public Function PKEY_Music_ArtistSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDEEB2DB5, &H696, &H4CE0, &H94, &HFE, &HA0, &H1F, &H77, &HA4, &H5F, &HB5, 102)
PKEY_Music_ArtistSortOverride = pkk
End Function
Public Function PKEY_Music_BeatsPerMinute() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 35)
PKEY_Music_BeatsPerMinute = pkk
End Function
Public Function PKEY_Music_Composer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 19)
PKEY_Music_Composer = pkk
End Function
Public Function PKEY_Music_ComposerSortOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBC20A3, &HBD48, &H4085, &H87, &H2C, &HA8, &H8D, &H77, &HF5, &H9, &H7E, 105)
PKEY_Music_ComposerSortOverride = pkk
End Function
Public Function PKEY_Music_Conductor() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 36)
PKEY_Music_Conductor = pkk
End Function
Public Function PKEY_Music_ContentGroupDescription() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 33)
PKEY_Music_ContentGroupDescription = pkk
End Function
Public Function PKEY_Music_DisplayArtist() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFD122953, &HFA93, &H4EF7, &H92, &HC3, &H4, &HC9, &H46, &HB2, &HF7, &HC8, 100)
PKEY_Music_DisplayArtist = pkk
End Function
Public Function PKEY_Music_Genre() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 11)
PKEY_Music_Genre = pkk
End Function
Public Function PKEY_Music_InitialKey() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 34)
PKEY_Music_InitialKey = pkk
End Function
Public Function PKEY_Music_IsCompilation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC449D5CB, &H9EA4, &H4809, &H82, &HE8, &HAF, &H9D, &H59, &HDE, &HD6, &HD1, 100)
PKEY_Music_IsCompilation = pkk
End Function
Public Function PKEY_Music_Lyrics() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 12)
PKEY_Music_Lyrics = pkk
End Function
Public Function PKEY_Music_Mood() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 39)
PKEY_Music_Mood = pkk
End Function
Public Function PKEY_Music_PartOfSet() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 37)
PKEY_Music_PartOfSet = pkk
End Function
Public Function PKEY_Music_Period() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 31)
PKEY_Music_Period = pkk
End Function
Public Function PKEY_Music_SynchronizedLyrics() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6B223B6A, &H162E, &H4AA9, &HB3, &H9F, &H5, &HD6, &H78, &HFC, &H6D, &H77, 100)
PKEY_Music_SynchronizedLyrics = pkk
End Function
Public Function PKEY_Music_TrackNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H56A3372E, &HCE9C, &H11D2, &H9F, &HE, &H0, &H60, &H97, &HC6, &H86, &HF6, 7)
PKEY_Music_TrackNumber = pkk
End Function
Public Function PKEY_Note_Color() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4776CAFA, &HBCE4, &H4CB1, &HA2, &H3E, &H26, &H5E, &H76, &HD8, &HEB, &H11, 100)
PKEY_Note_Color = pkk
End Function
Public Function PKEY_Note_ColorText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H46B4E8DE, &HCDB2, &H440D, &H88, &H5C, &H16, &H58, &HEB, &H65, &HB9, &H14, 100)
PKEY_Note_ColorText = pkk
End Function
Public Function PKEY_Photo_Aperture() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37378)
PKEY_Photo_Aperture = pkk
End Function
Public Function PKEY_Photo_ApertureDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE1A9A38B, &H6685, &H46BD, &H87, &H5E, &H57, &HD, &HC7, &HAD, &H73, &H20, 100)
PKEY_Photo_ApertureDenominator = pkk
End Function
Public Function PKEY_Photo_ApertureNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H337ECEC, &H39FB, &H4581, &HA0, &HBD, &H4C, &H4C, &HC5, &H1E, &H99, &H14, 100)
PKEY_Photo_ApertureNumerator = pkk
End Function
Public Function PKEY_Photo_Brightness() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1A701BF6, &H478C, &H4361, &H83, &HAB, &H37, &H1, &HBB, &H5, &H3C, &H58, 100)
PKEY_Photo_Brightness = pkk
End Function
Public Function PKEY_Photo_BrightnessDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6EBE6946, &H2321, &H440A, &H90, &HF0, &HC0, &H43, &HEF, &HD3, &H24, &H76, 100)
PKEY_Photo_BrightnessDenominator = pkk
End Function
Public Function PKEY_Photo_BrightnessNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9E7D118F, &HB314, &H45A0, &H8C, &HFB, &HD6, &H54, &HB9, &H17, &HC9, &HE9, 100)
PKEY_Photo_BrightnessNumerator = pkk
End Function
Public Function PKEY_Photo_CameraManufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 271)
PKEY_Photo_CameraManufacturer = pkk
End Function
Public Function PKEY_Photo_CameraModel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 272)
PKEY_Photo_CameraModel = pkk
End Function
Public Function PKEY_Photo_CameraSerialNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 273)
PKEY_Photo_CameraSerialNumber = pkk
End Function
Public Function PKEY_Photo_Contrast() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2A785BA9, &H8D23, &H4DED, &H82, &HE6, &H60, &HA3, &H50, &HC8, &H6A, &H10, 100)
PKEY_Photo_Contrast = pkk
End Function
Public Function PKEY_Photo_ContrastText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H59DDE9F2, &H5253, &H40EA, &H9A, &H8B, &H47, &H9E, &H96, &HC6, &H24, &H9A, 100)
PKEY_Photo_ContrastText = pkk
End Function
Public Function PKEY_Photo_DateTaken() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 36867)
PKEY_Photo_DateTaken = pkk
End Function
Public Function PKEY_Photo_DigitalZoom() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF85BF840, &HA925, &H4BC2, &HB0, &HC4, &H8E, &H36, &HB5, &H98, &H67, &H9E, 100)
PKEY_Photo_DigitalZoom = pkk
End Function
Public Function PKEY_Photo_DigitalZoomDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H745BAF0E, &HE5C1, &H4CFB, &H8A, &H1B, &HD0, &H31, &HA0, &HA5, &H23, &H93, 100)
PKEY_Photo_DigitalZoomDenominator = pkk
End Function
Public Function PKEY_Photo_DigitalZoomNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H16CBB924, &H6500, &H473B, &HA5, &HBE, &HF1, &H59, &H9B, &HCB, &HE4, &H13, 100)
PKEY_Photo_DigitalZoomNumerator = pkk
End Function
Public Function PKEY_Photo_Event() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 18248)
PKEY_Photo_Event = pkk
End Function
Public Function PKEY_Photo_EXIFVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD35F743A, &HEB2E, &H47F2, &HA2, &H86, &H84, &H41, &H32, &HCB, &H14, &H27, 100)
PKEY_Photo_EXIFVersion = pkk
End Function
Public Function PKEY_Photo_ExposureBias() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37380)
PKEY_Photo_ExposureBias = pkk
End Function
Public Function PKEY_Photo_ExposureBiasDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAB205E50, &H4B7, &H461C, &HA1, &H8C, &H2F, &H23, &H38, &H36, &HE6, &H27, 100)
PKEY_Photo_ExposureBiasDenominator = pkk
End Function
Public Function PKEY_Photo_ExposureBiasNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H738BF284, &H1D87, &H420B, &H92, &HCF, &H58, &H34, &HBF, &H6E, &HF9, &HED, 100)
PKEY_Photo_ExposureBiasNumerator = pkk
End Function
Public Function PKEY_Photo_ExposureIndex() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H967B5AF8, &H995A, &H46ED, &H9E, &H11, &H35, &HB3, &HC5, &HB9, &H78, &H2D, 100)
PKEY_Photo_ExposureIndex = pkk
End Function
Public Function PKEY_Photo_ExposureIndexDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H93112F89, &HC28B, &H492F, &H8A, &H9D, &H4B, &HE2, &H6, &H2C, &HEE, &H8A, 100)
PKEY_Photo_ExposureIndexDenominator = pkk
End Function
Public Function PKEY_Photo_ExposureIndexNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCDEDCF30, &H8919, &H44DF, &H8F, &H4C, &H4E, &HB2, &HFF, &HDB, &H8D, &H89, 100)
PKEY_Photo_ExposureIndexNumerator = pkk
End Function
Public Function PKEY_Photo_ExposureProgram() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 34850)
PKEY_Photo_ExposureProgram = pkk
End Function
Public Function PKEY_Photo_ExposureProgramText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFEC690B7, &H5F30, &H4646, &HAE, &H47, &H4C, &HAA, &HFB, &HA8, &H84, &HA3, 100)
PKEY_Photo_ExposureProgramText = pkk
End Function
Public Function PKEY_Photo_ExposureTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 33434)
PKEY_Photo_ExposureTime = pkk
End Function
Public Function PKEY_Photo_ExposureTimeDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H55E98597, &HAD16, &H42E0, &HB6, &H24, &H21, &H59, &H9A, &H19, &H98, &H38, 100)
PKEY_Photo_ExposureTimeDenominator = pkk
End Function
Public Function PKEY_Photo_ExposureTimeNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H257E44E2, &H9031, &H4323, &HAC, &H38, &H85, &HC5, &H52, &H87, &H1B, &H2E, 100)
PKEY_Photo_ExposureTimeNumerator = pkk
End Function
Public Function PKEY_Photo_Flash() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37385)
PKEY_Photo_Flash = pkk
End Function
Public Function PKEY_Photo_FlashEnergy() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 41483)
PKEY_Photo_FlashEnergy = pkk
End Function
Public Function PKEY_Photo_FlashEnergyDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD7B61C70, &H6323, &H49CD, &HA5, &HFC, &HC8, &H42, &H77, &H16, &H2C, &H97, 100)
PKEY_Photo_FlashEnergyDenominator = pkk
End Function
Public Function PKEY_Photo_FlashEnergyNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFCAD3D3D, &H858, &H400F, &HAA, &HA3, &H2F, &H66, &HCC, &HE2, &HA6, &HBC, 100)
PKEY_Photo_FlashEnergyNumerator = pkk
End Function
Public Function PKEY_Photo_FlashManufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAABAF6C9, &HE0C5, &H4719, &H85, &H85, &H57, &HB1, &H3, &HE5, &H84, &HFE, 100)
PKEY_Photo_FlashManufacturer = pkk
End Function
Public Function PKEY_Photo_FlashModel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFE83BB35, &H4D1A, &H42E2, &H91, &H6B, &H6, &HF3, &HE1, &HAF, &H71, &H9E, 100)
PKEY_Photo_FlashModel = pkk
End Function
Public Function PKEY_Photo_FlashText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6B8B68F6, &H200B, &H47EA, &H8D, &H25, &HD8, &H5, &HF, &H57, &H33, &H9F, 100)
PKEY_Photo_FlashText = pkk
End Function
Public Function PKEY_Photo_FNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 33437)
PKEY_Photo_FNumber = pkk
End Function
Public Function PKEY_Photo_FNumberDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE92A2496, &H223B, &H4463, &HA4, &HE3, &H30, &HEA, &HBB, &HA7, &H9D, &H80, 100)
PKEY_Photo_FNumberDenominator = pkk
End Function
Public Function PKEY_Photo_FNumberNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1B97738A, &HFDFC, &H462F, &H9D, &H93, &H19, &H57, &HE0, &H8B, &HE9, &HC, 100)
PKEY_Photo_FNumberNumerator = pkk
End Function
Public Function PKEY_Photo_FocalLength() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37386)
PKEY_Photo_FocalLength = pkk
End Function
Public Function PKEY_Photo_FocalLengthDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H305BC615, &HDCA1, &H44A5, &H9F, &HD4, &H10, &HC0, &HBA, &H79, &H41, &H2E, 100)
PKEY_Photo_FocalLengthDenominator = pkk
End Function
Public Function PKEY_Photo_FocalLengthInFilm() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA0E74609, &HB84D, &H4F49, &HB8, &H60, &H46, &H2B, &HD9, &H97, &H1F, &H98, 100)
PKEY_Photo_FocalLengthInFilm = pkk
End Function
Public Function PKEY_Photo_FocalLengthNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H776B6B3B, &H1E3D, &H4B0C, &H9A, &HE, &H8F, &HBA, &HF2, &HA8, &H49, &H2A, 100)
PKEY_Photo_FocalLengthNumerator = pkk
End Function
Public Function PKEY_Photo_FocalPlaneXResolution() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCFC08D97, &HC6F7, &H4484, &H89, &HDD, &HEB, &HEF, &H43, &H56, &HFE, &H76, 100)
PKEY_Photo_FocalPlaneXResolution = pkk
End Function
Public Function PKEY_Photo_FocalPlaneXResolutionDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H933F3F5, &H4786, &H4F46, &HA8, &HE8, &HD6, &H4D, &HD3, &H7F, &HA5, &H21, 100)
PKEY_Photo_FocalPlaneXResolutionDenominator = pkk
End Function
Public Function PKEY_Photo_FocalPlaneXResolutionNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDCCB10AF, &HB4E2, &H4B88, &H95, &HF9, &H3, &H1B, &H4D, &H5A, &HB4, &H90, 100)
PKEY_Photo_FocalPlaneXResolutionNumerator = pkk
End Function
Public Function PKEY_Photo_FocalPlaneYResolution() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4FFFE4D0, &H914F, &H4AC4, &H8D, &H6F, &HC9, &HC6, &H1D, &HE1, &H69, &HB1, 100)
PKEY_Photo_FocalPlaneYResolution = pkk
End Function
Public Function PKEY_Photo_FocalPlaneYResolutionDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1D6179A6, &HA876, &H4031, &HB0, &H13, &H33, &H47, &HB2, &HB6, &H4D, &HC8, 100)
PKEY_Photo_FocalPlaneYResolutionDenominator = pkk
End Function
Public Function PKEY_Photo_FocalPlaneYResolutionNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA2E541C5, &H4440, &H4BA8, &H86, &H7E, &H75, &HCF, &HC0, &H68, &H28, &HCD, 100)
PKEY_Photo_FocalPlaneYResolutionNumerator = pkk
End Function
Public Function PKEY_Photo_GainControl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFA304789, &HC7, &H4D80, &H90, &H4A, &H1E, &H4D, &HCC, &H72, &H65, &HAA, 100)
PKEY_Photo_GainControl = pkk
End Function
Public Function PKEY_Photo_GainControlDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H42864DFD, &H9DA4, &H4F77, &HBD, &HED, &H4A, &HAD, &H7B, &H25, &H67, &H35, 100)
PKEY_Photo_GainControlDenominator = pkk
End Function
Public Function PKEY_Photo_GainControlNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8E8ECF7C, &HB7B8, &H4EB8, &HA6, &H3F, &HE, &HE7, &H15, &HC9, &H6F, &H9E, 100)
PKEY_Photo_GainControlNumerator = pkk
End Function
Public Function PKEY_Photo_GainControlText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC06238B2, &HBF9, &H4279, &HA7, &H23, &H25, &H85, &H67, &H15, &HCB, &H9D, 100)
PKEY_Photo_GainControlText = pkk
End Function
Public Function PKEY_Photo_ISOSpeed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 34855)
PKEY_Photo_ISOSpeed = pkk
End Function
Public Function PKEY_Photo_LensManufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE6DDCAF7, &H29C5, &H4F0A, &H9A, &H68, &HD1, &H94, &H12, &HEC, &H70, &H90, 100)
PKEY_Photo_LensManufacturer = pkk
End Function
Public Function PKEY_Photo_LensModel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE1277516, &H2B5F, &H4869, &H89, &HB1, &H2E, &H58, &H5B, &HD3, &H8B, &H7A, 100)
PKEY_Photo_LensModel = pkk
End Function
Public Function PKEY_Photo_LightSource() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37384)
PKEY_Photo_LightSource = pkk
End Function
Public Function PKEY_Photo_MakerNote() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFA303353, &HB659, &H4052, &H85, &HE9, &HBC, &HAC, &H79, &H54, &H9B, &H84, 100)
PKEY_Photo_MakerNote = pkk
End Function
Public Function PKEY_Photo_MakerNoteOffset() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H813F4124, &H34E6, &H4D17, &HAB, &H3E, &H6B, &H1F, &H3C, &H22, &H47, &HA1, 100)
PKEY_Photo_MakerNoteOffset = pkk
End Function
Public Function PKEY_Photo_MaxAperture() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8F6D7C2, &HE3F2, &H44FC, &HAF, &H1E, &H5A, &HA5, &HC8, &H1A, &H2D, &H3E, 100)
PKEY_Photo_MaxAperture = pkk
End Function
Public Function PKEY_Photo_MaxApertureDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC77724D4, &H601F, &H46C5, &H9B, &H89, &HC5, &H3F, &H93, &HBC, &HEB, &H77, 100)
PKEY_Photo_MaxApertureDenominator = pkk
End Function
Public Function PKEY_Photo_MaxApertureNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC107E191, &HA459, &H44C5, &H9A, &HE6, &HB9, &H52, &HAD, &H4B, &H90, &H6D, 100)
PKEY_Photo_MaxApertureNumerator = pkk
End Function
Public Function PKEY_Photo_MeteringMode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37383)
PKEY_Photo_MeteringMode = pkk
End Function
Public Function PKEY_Photo_MeteringModeText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF628FD8C, &H7BA8, &H465A, &HA6, &H5B, &HC5, &HAA, &H79, &H26, &H3A, &H9E, 100)
PKEY_Photo_MeteringModeText = pkk
End Function
Public Function PKEY_Photo_Orientation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 274)
PKEY_Photo_Orientation = pkk
End Function
Public Function PKEY_Photo_OrientationText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA9EA193C, &HC511, &H498A, &HA0, &H6B, &H58, &HE2, &H77, &H6D, &HCC, &H28, 100)
PKEY_Photo_OrientationText = pkk
End Function
Public Function PKEY_Photo_PeopleNames() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE8309B6E, &H84C, &H49B4, &HB1, &HFC, &H90, &HA8, &H3, &H31, &HB6, &H38, 100)
PKEY_Photo_PeopleNames = pkk
End Function
Public Function PKEY_Photo_PhotometricInterpretation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H341796F1, &H1DF9, &H4B1C, &HA5, &H64, &H91, &HBD, &HEF, &HA4, &H38, &H77, 100)
PKEY_Photo_PhotometricInterpretation = pkk
End Function
Public Function PKEY_Photo_PhotometricInterpretationText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H821437D6, &H9EAB, &H4765, &HA5, &H89, &H3B, &H1C, &HBB, &HD2, &H2A, &H61, 100)
PKEY_Photo_PhotometricInterpretationText = pkk
End Function
Public Function PKEY_Photo_ProgramMode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D217F6D, &H3F6A, &H4825, &HB4, &H70, &H5F, &H3, &HCA, &H2F, &HBE, &H9B, 100)
PKEY_Photo_ProgramMode = pkk
End Function
Public Function PKEY_Photo_ProgramModeText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7FE3AA27, &H2648, &H42F3, &H89, &HB0, &H45, &H4E, &H5C, &HB1, &H50, &HC3, 100)
PKEY_Photo_ProgramModeText = pkk
End Function
Public Function PKEY_Photo_RelatedSoundFile() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H318A6B45, &H87F, &H4DC2, &HB8, &HCC, &H5, &H35, &H95, &H51, &HFC, &H9E, 100)
PKEY_Photo_RelatedSoundFile = pkk
End Function
Public Function PKEY_Photo_Saturation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49237325, &HA95A, &H4F67, &HB2, &H11, &H81, &H6B, &H2D, &H45, &HD2, &HE0, 100)
PKEY_Photo_Saturation = pkk
End Function
Public Function PKEY_Photo_SaturationText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H61478C08, &HB600, &H4A84, &HBB, &HE4, &HE9, &H9C, &H45, &HF0, &HA0, &H72, 100)
PKEY_Photo_SaturationText = pkk
End Function
Public Function PKEY_Photo_Sharpness() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HFC6976DB, &H8349, &H4970, &HAE, &H97, &HB3, &HC5, &H31, &H6A, &H8, &HF0, 100)
PKEY_Photo_Sharpness = pkk
End Function
Public Function PKEY_Photo_SharpnessText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H51EC3F47, &HDD50, &H421D, &H87, &H69, &H33, &H4F, &H50, &H42, &H4B, &H1E, 100)
PKEY_Photo_SharpnessText = pkk
End Function
Public Function PKEY_Photo_ShutterSpeed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37377)
PKEY_Photo_ShutterSpeed = pkk
End Function
Public Function PKEY_Photo_ShutterSpeedDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE13D8975, &H81C7, &H4948, &HAE, &H3F, &H37, &HCA, &HE1, &H1E, &H8F, &HF7, 100)
PKEY_Photo_ShutterSpeedDenominator = pkk
End Function
Public Function PKEY_Photo_ShutterSpeedNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H16EA4042, &HD6F4, &H4BCA, &H83, &H49, &H7C, &H78, &HD3, &HF, &HB3, &H33, 100)
PKEY_Photo_ShutterSpeedNumerator = pkk
End Function
Public Function PKEY_Photo_SubjectDistance() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14B81DA1, &H135, &H4D31, &H96, &HD9, &H6C, &HBF, &HC9, &H67, &H1A, &H99, 37382)
PKEY_Photo_SubjectDistance = pkk
End Function
Public Function PKEY_Photo_SubjectDistanceDenominator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC840A88, &HB043, &H466D, &H97, &H66, &HD4, &HB2, &H6D, &HA3, &HFA, &H77, 100)
PKEY_Photo_SubjectDistanceDenominator = pkk
End Function
Public Function PKEY_Photo_SubjectDistanceNumerator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8AF4961C, &HF526, &H43E5, &HAA, &H81, &HDB, &H76, &H82, &H19, &H17, &H8D, 100)
PKEY_Photo_SubjectDistanceNumerator = pkk
End Function
Public Function PKEY_Photo_TagViewAggregate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB812F15D, &HC2D8, &H4BBF, &HBA, &HCD, &H79, &H74, &H43, &H46, &H11, &H3F, 100)
PKEY_Photo_TagViewAggregate = pkk
End Function
Public Function PKEY_Photo_TranscodedForSync() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9A8EBB75, &H6458, &H4E82, &HBA, &HCB, &H35, &HC0, &H9, &H5B, &H3, &HBB, 100)
PKEY_Photo_TranscodedForSync = pkk
End Function
Public Function PKEY_Photo_WhiteBalance() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HEE3D3D8A, &H5381, &H4CFA, &HB1, &H3B, &HAA, &HF6, &H6B, &H5F, &H4E, &HC9, 100)
PKEY_Photo_WhiteBalance = pkk
End Function
Public Function PKEY_Photo_WhiteBalanceText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6336B95E, &HC7A7, &H426D, &H86, &HFD, &H7A, &HE3, &HD3, &H9C, &H84, &HB4, 100)
PKEY_Photo_WhiteBalanceText = pkk
End Function
Public Function PKEY_PropGroup_Advanced() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H900A403B, &H97B, &H4B95, &H8A, &HE2, &H7, &H1F, &HDA, &HEE, &HB1, &H18, 100)
PKEY_PropGroup_Advanced = pkk
End Function
Public Function PKEY_PropGroup_Audio() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2804D469, &H788F, &H48AA, &H85, &H70, &H71, &HB9, &HC1, &H87, &HE1, &H38, 100)
PKEY_PropGroup_Audio = pkk
End Function
Public Function PKEY_PropGroup_Calendar() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9973D2B5, &HBFD8, &H438A, &HBA, &H94, &H53, &H49, &HB2, &H93, &H18, &H1A, 100)
PKEY_PropGroup_Calendar = pkk
End Function
Public Function PKEY_PropGroup_Camera() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDE00DE32, &H547E, &H4981, &HAD, &H4B, &H54, &H2F, &H2E, &H90, &H7, &HD8, 100)
PKEY_PropGroup_Camera = pkk
End Function
Public Function PKEY_PropGroup_Contact() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HDF975FD3, &H250A, &H4004, &H85, &H8F, &H34, &HE2, &H9A, &H3E, &H37, &HAA, 100)
PKEY_PropGroup_Contact = pkk
End Function
Public Function PKEY_PropGroup_Content() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD0DAB0BA, &H368A, &H4050, &HA8, &H82, &H6C, &H1, &HF, &HD1, &H9A, &H4F, 100)
PKEY_PropGroup_Content = pkk
End Function
Public Function PKEY_PropGroup_Description() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8969B275, &H9475, &H4E00, &HA8, &H87, &HFF, &H93, &HB8, &HB4, &H1E, &H44, 100)
PKEY_PropGroup_Description = pkk
End Function
Public Function PKEY_PropGroup_FileSystem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3A7D2C1, &H80FC, &H4B40, &H8F, &H34, &H30, &HEA, &H11, &H1B, &HDC, &H2E, 100)
PKEY_PropGroup_FileSystem = pkk
End Function
Public Function PKEY_PropGroup_General() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCC301630, &HB192, &H4C22, &HB3, &H72, &H9F, &H4C, &H6D, &H33, &H8E, &H7, 100)
PKEY_PropGroup_General = pkk
End Function
Public Function PKEY_PropGroup_GPS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF3713ADA, &H90E3, &H4E11, &HAA, &HE5, &HFD, &HC1, &H76, &H85, &HB9, &HBE, 100)
PKEY_PropGroup_GPS = pkk
End Function
Public Function PKEY_PropGroup_Image() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE3690A87, &HFA8, &H4A2A, &H9A, &H9F, &HFC, &HE8, &H82, &H70, &H55, &HAC, 100)
PKEY_PropGroup_Image = pkk
End Function
Public Function PKEY_PropGroup_Media() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H61872CF7, &H6B5E, &H4B4B, &HAC, &H2D, &H59, &HDA, &H84, &H45, &H92, &H48, 100)
PKEY_PropGroup_Media = pkk
End Function
Public Function PKEY_PropGroup_MediaAdvanced() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8859A284, &HDE7E, &H4642, &H99, &HBA, &HD4, &H31, &HD0, &H44, &HB1, &HEC, 100)
PKEY_PropGroup_MediaAdvanced = pkk
End Function
Public Function PKEY_PropGroup_Message() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7FD7259D, &H16B4, &H4135, &H9F, &H97, &H7C, &H96, &HEC, &HD2, &HFA, &H9E, 100)
PKEY_PropGroup_Message = pkk
End Function
Public Function PKEY_PropGroup_Music() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H68DD6094, &H7216, &H40F1, &HA0, &H29, &H43, &HFE, &H71, &H27, &H4, &H3F, 100)
PKEY_PropGroup_Music = pkk
End Function
Public Function PKEY_PropGroup_Origin() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2598D2FB, &H5569, &H4367, &H95, &HDF, &H5C, &HD3, &HA1, &H77, &HE1, &HA5, 100)
PKEY_PropGroup_Origin = pkk
End Function
Public Function PKEY_PropGroup_PhotoAdvanced() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCB2BF5A, &H9EE7, &H4A86, &H82, &H22, &HF0, &H1E, &H7, &HFD, &HAD, &HAF, 100)
PKEY_PropGroup_PhotoAdvanced = pkk
End Function
Public Function PKEY_PropGroup_RecordedTV() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HE7B33238, &H6584, &H4170, &HA5, &HC0, &HAC, &H25, &HEF, &HD9, &HDA, &H56, 100)
PKEY_PropGroup_RecordedTV = pkk
End Function
Public Function PKEY_PropGroup_Video() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBEBE0920, &H7671, &H4C54, &HA3, &HEB, &H49, &HFD, &HDF, &HC1, &H91, &HEE, 100)
PKEY_PropGroup_Video = pkk
End Function
Public Function PKEY_InfoTipText() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 17)
PKEY_InfoTipText = pkk
End Function
Public Function PKEY_PropList_ConflictPrompt() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 11)
PKEY_PropList_ConflictPrompt = pkk
End Function
Public Function PKEY_PropList_ContentViewModeForBrowse() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 13)
PKEY_PropList_ContentViewModeForBrowse = pkk
End Function
Public Function PKEY_PropList_ContentViewModeForSearch() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 14)
PKEY_PropList_ContentViewModeForSearch = pkk
End Function
Public Function PKEY_PropList_ExtendedTileInfo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 9)
PKEY_PropList_ExtendedTileInfo = pkk
End Function
Public Function PKEY_PropList_FileOperationPrompt() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 10)
PKEY_PropList_FileOperationPrompt = pkk
End Function
Public Function PKEY_PropList_FullDetails() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 2)
PKEY_PropList_FullDetails = pkk
End Function
Public Function PKEY_PropList_InfoTip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 4)
PKEY_PropList_InfoTip = pkk
End Function
Public Function PKEY_PropList_NonPersonal() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49D1091F, &H82E, &H493F, &HB2, &H3F, &HD2, &H30, &H8A, &HA9, &H66, &H8C, 100)
PKEY_PropList_NonPersonal = pkk
End Function
Public Function PKEY_PropList_PreviewDetails() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 8)
PKEY_PropList_PreviewDetails = pkk
End Function
Public Function PKEY_PropList_PreviewTitle() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 6)
PKEY_PropList_PreviewTitle = pkk
End Function
Public Function PKEY_PropList_QuickTip() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 5)
PKEY_PropList_QuickTip = pkk
End Function
Public Function PKEY_PropList_TileInfo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HC9944A21, &HA406, &H48FE, &H82, &H25, &HAE, &HC7, &HE2, &H4C, &H21, &H1B, 3)
PKEY_PropList_TileInfo = pkk
End Function
Public Function PKEY_PropList_XPDetailsPanel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HF2275480, &HF782, &H4291, &HBD, &H94, &HF1, &H36, &H93, &H51, &H3A, &HEC, 0)
PKEY_PropList_XPDetailsPanel = pkk
End Function
Public Function PKEY_RecordedTV_ChannelNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 7)
PKEY_RecordedTV_ChannelNumber = pkk
End Function
Public Function PKEY_RecordedTV_Credits() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 4)
PKEY_RecordedTV_Credits = pkk
End Function
Public Function PKEY_RecordedTV_DateContentExpires() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 15)
PKEY_RecordedTV_DateContentExpires = pkk
End Function
Public Function PKEY_RecordedTV_EpisodeName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 2)
PKEY_RecordedTV_EpisodeName = pkk
End Function
Public Function PKEY_RecordedTV_IsATSCContent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 16)
PKEY_RecordedTV_IsATSCContent = pkk
End Function
Public Function PKEY_RecordedTV_IsClosedCaptioningAvailable() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 12)
PKEY_RecordedTV_IsClosedCaptioningAvailable = pkk
End Function
Public Function PKEY_RecordedTV_IsDTVContent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 17)
PKEY_RecordedTV_IsDTVContent = pkk
End Function
Public Function PKEY_RecordedTV_IsHDContent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 18)
PKEY_RecordedTV_IsHDContent = pkk
End Function
Public Function PKEY_RecordedTV_IsRepeatBroadcast() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 13)
PKEY_RecordedTV_IsRepeatBroadcast = pkk
End Function
Public Function PKEY_RecordedTV_IsSAP() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 14)
PKEY_RecordedTV_IsSAP = pkk
End Function
Public Function PKEY_RecordedTV_NetworkAffiliation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H2C53C813, &HFB63, &H4E22, &HA1, &HAB, &HB, &H33, &H1C, &HA1, &HE2, &H73, 100)
PKEY_RecordedTV_NetworkAffiliation = pkk
End Function
Public Function PKEY_RecordedTV_OriginalBroadcastDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4684FE97, &H8765, &H4842, &H9C, &H13, &HF0, &H6, &H44, &H7B, &H17, &H8C, 100)
PKEY_RecordedTV_OriginalBroadcastDate = pkk
End Function
Public Function PKEY_RecordedTV_ProgramDescription() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 3)
PKEY_RecordedTV_ProgramDescription = pkk
End Function
Public Function PKEY_RecordedTV_RecordingTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA5477F61, &H7A82, &H4ECA, &H9D, &HDE, &H98, &HB6, &H9B, &H24, &H79, &HB3, 100)
PKEY_RecordedTV_RecordingTime = pkk
End Function
Public Function PKEY_RecordedTV_StationCallSign() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H6D748DE2, &H8D38, &H4CC3, &HAC, &H60, &HF0, &H9, &HB0, &H57, &HC5, &H57, 5)
PKEY_RecordedTV_StationCallSign = pkk
End Function
Public Function PKEY_RecordedTV_StationName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H1B5439E7, &HEBA1, &H4AF8, &HBD, &HD7, &H7A, &HF1, &HD4, &H54, &H94, &H93, 100)
PKEY_RecordedTV_StationName = pkk
End Function
Public Function PKEY_Search_AutoSummary() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H560C36C0, &H503A, &H11CF, &HBA, &HA1, &H0, &H0, &H4C, &H75, &H2A, &H9A, 2)
PKEY_Search_AutoSummary = pkk
End Function
Public Function PKEY_Search_ContainerHash() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HBCEEE283, &H35DF, &H4D53, &H82, &H6A, &HF3, &H6A, &H3E, &HEF, &HC6, &HBE, 100)
PKEY_Search_ContainerHash = pkk
End Function
Public Function PKEY_Search_Contents() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 19)
PKEY_Search_Contents = pkk
End Function
Public Function PKEY_Search_EntryID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49691C90, &H7E17, &H101A, &HA9, &H1C, &H8, &H0, &H2B, &H2E, &HCD, &HA9, 5)
PKEY_Search_EntryID = pkk
End Function
Public Function PKEY_Search_ExtendedProperties() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7B03B546, &HFA4F, &H4A52, &HA2, &HFE, &H3, &HD5, &H31, &H1E, &H58, &H65, 100)
PKEY_Search_ExtendedProperties = pkk
End Function
Public Function PKEY_Search_GatherTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E350, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 8)
PKEY_Search_GatherTime = pkk
End Function
Public Function PKEY_Search_HitCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49691C90, &H7E17, &H101A, &HA9, &H1C, &H8, &H0, &H2B, &H2E, &HCD, &HA9, 4)
PKEY_Search_HitCount = pkk
End Function
Public Function PKEY_Search_IsClosedDirectory() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E343, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 23)
PKEY_Search_IsClosedDirectory = pkk
End Function
Public Function PKEY_Search_IsFullyContained() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E343, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 24)
PKEY_Search_IsFullyContained = pkk
End Function
Public Function PKEY_Search_QueryFocusedSummary() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H560C36C0, &H503A, &H11CF, &HBA, &HA1, &H0, &H0, &H4C, &H75, &H2A, &H9A, 3)
PKEY_Search_QueryFocusedSummary = pkk
End Function
Public Function PKEY_Search_QueryFocusedSummaryWithFallback() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H560C36C0, &H503A, &H11CF, &HBA, &HA1, &H0, &H0, &H4C, &H75, &H2A, &H9A, 4)
PKEY_Search_QueryFocusedSummaryWithFallback = pkk
End Function
Public Function PKEY_Search_QueryPropertyHits() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49691C90, &H7E17, &H101A, &HA9, &H1C, &H8, &H0, &H2B, &H2E, &HCD, &HA9, 21)
PKEY_Search_QueryPropertyHits = pkk
End Function
Public Function PKEY_Search_Rank() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H49691C90, &H7E17, &H101A, &HA9, &H1C, &H8, &H0, &H2B, &H2E, &HCD, &HA9, 3)
PKEY_Search_Rank = pkk
End Function
Public Function PKEY_Search_Store() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA06992B3, &H8CAF, &H4ED7, &HA5, &H47, &HB2, &H59, &HE3, &H2A, &HC9, &HFC, 100)
PKEY_Search_Store = pkk
End Function
Public Function PKEY_Search_UrlToIndex() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E343, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 2)
PKEY_Search_UrlToIndex = pkk
End Function
Public Function PKEY_Search_UrlToIndexWithModificationTime() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB63E343, &H9CCC, &H11D0, &HBC, &HDB, &H0, &H80, &H5F, &HCC, &HCE, &H4, 12)
PKEY_Search_UrlToIndexWithModificationTime = pkk
End Function
Public Function PKEY_DescriptionID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 2)
PKEY_DescriptionID = pkk
End Function
Public Function PKEY_InternalName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 5)
PKEY_InternalName = pkk
End Function
Public Function PKEY_LibraryLocationsCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H908696C7, &H8F87, &H44F2, &H80, &HED, &HA8, &HC1, &HC6, &H89, &H45, &H75, 2)
PKEY_LibraryLocationsCount = pkk
End Function
Public Function PKEY_Link_TargetSFGAOFlagsStrings() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD6942081, &HD53B, &H443D, &HAD, &H47, &H5E, &H5, &H9D, &H9C, &HD2, &H7A, 3)
PKEY_Link_TargetSFGAOFlagsStrings = pkk
End Function
Public Function PKEY_Link_TargetUrl() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H5CBF2787, &H48CF, &H4208, &HB9, &HE, &HEE, &H5E, &H5D, &H42, &H2, &H94, 2)
PKEY_Link_TargetUrl = pkk
End Function
Public Function PKEY_NamespaceCLSID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H28636AA6, &H953D, &H11D2, &HB5, &HD6, &H0, &HC0, &H4F, &HD9, &H18, &HD0, 6)
PKEY_NamespaceCLSID = pkk
End Function
Public Function PKEY_Shell_SFGAOFlagsStrings() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD6942081, &HD53B, &H443D, &HAD, &H47, &H5E, &H5, &H9D, &H9C, &HD2, &H7A, 2)
PKEY_Shell_SFGAOFlagsStrings = pkk
End Function
Public Function PKEY_StatusBarSelectedItemCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26DC287C, &H6E3D, &H4BD3, &HB2, &HB0, &H6A, &H26, &HBA, &H2E, &H34, &H6D, 3)
PKEY_StatusBarSelectedItemCount = pkk
End Function
Public Function PKEY_StatusBarViewItemCount() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26DC287C, &H6E3D, &H4BD3, &HB2, &HB0, &H6A, &H26, &HBA, &H2E, &H34, &H6D, 2)
PKEY_StatusBarViewItemCount = pkk
End Function
Public Function PKEY_AppUserModel_ExcludeFromShowInNewInstall() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 8)
PKEY_AppUserModel_ExcludeFromShowInNewInstall = pkk
End Function
Public Function PKEY_AppUserModel_ID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 5)
PKEY_AppUserModel_ID = pkk
End Function
Public Function PKEY_AppUserModel_IsDestListSeparator() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 6)
PKEY_AppUserModel_IsDestListSeparator = pkk
End Function
Public Function PKEY_AppUserModel_IsDualMode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 11)
PKEY_AppUserModel_IsDualMode = pkk
End Function
Public Function PKEY_AppUserModel_PreventPinning() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 9)
PKEY_AppUserModel_PreventPinning = pkk
End Function
Public Function PKEY_AppUserModel_RelaunchCommand() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 2)
PKEY_AppUserModel_RelaunchCommand = pkk
End Function
Public Function PKEY_AppUserModel_RelaunchDisplayNameResource() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 4)
PKEY_AppUserModel_RelaunchDisplayNameResource = pkk
End Function
Public Function PKEY_AppUserModel_RelaunchIconResource() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 3)
PKEY_AppUserModel_RelaunchIconResource = pkk
End Function
Public Function PKEY_AppUserModel_StartPinOption() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 12)
PKEY_AppUserModel_StartPinOption = pkk
End Function
Public Function PKEY_AppUserModel_ToastActivatorCLSID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9F4C2855, &H9F79, &H4B39, &HA8, &HD0, &HE1, &HD4, &H2D, &HE1, &HD5, &HF3, 26)
PKEY_AppUserModel_ToastActivatorCLSID = pkk
End Function
Public Function PKEY_EdgeGesture_DisableTouchWhenFullscreen() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H32CE38B2, &H2C9A, &H41B1, &H9B, &HC5, &HB3, &H78, &H43, &H94, &HAA, &H44, 2)
PKEY_EdgeGesture_DisableTouchWhenFullscreen = pkk
End Function
Public Function PKEY_Software_DateLastUsed() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H841E4F90, &HFF59, &H4D16, &H89, &H47, &HE8, &H1B, &HBF, &HFA, &HB3, &H6D, 16)
PKEY_Software_DateLastUsed = pkk
End Function
Public Function PKEY_Software_ProductName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCEF7D53, &HFA64, &H11D1, &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 7)
PKEY_Software_ProductName = pkk
End Function
Public Function PKEY_Sync_Comments() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 13)
PKEY_Sync_Comments = pkk
End Function
Public Function PKEY_Sync_ConflictDescription() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCE50C159, &H2FB8, &H41FD, &HBE, &H68, &HD3, &HE0, &H42, &HE2, &H74, &HBC, 4)
PKEY_Sync_ConflictDescription = pkk
End Function
Public Function PKEY_Sync_ConflictFirstLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCE50C159, &H2FB8, &H41FD, &HBE, &H68, &HD3, &HE0, &H42, &HE2, &H74, &HBC, 6)
PKEY_Sync_ConflictFirstLocation = pkk
End Function
Public Function PKEY_Sync_ConflictSecondLocation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCE50C159, &H2FB8, &H41FD, &HBE, &H68, &HD3, &HE0, &H42, &HE2, &H74, &HBC, 7)
PKEY_Sync_ConflictSecondLocation = pkk
End Function
Public Function PKEY_Sync_HandlerCollectionID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 2)
PKEY_Sync_HandlerCollectionID = pkk
End Function
Public Function PKEY_Sync_HandlerID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 3)
PKEY_Sync_HandlerID = pkk
End Function
Public Function PKEY_Sync_HandlerName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCE50C159, &H2FB8, &H41FD, &HBE, &H68, &HD3, &HE0, &H42, &HE2, &H74, &HBC, 2)
PKEY_Sync_HandlerName = pkk
End Function
Public Function PKEY_Sync_HandlerType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 8)
PKEY_Sync_HandlerType = pkk
End Function
Public Function PKEY_Sync_HandlerTypeLabel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 9)
PKEY_Sync_HandlerTypeLabel = pkk
End Function
Public Function PKEY_Sync_ItemID() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 6)
PKEY_Sync_ItemID = pkk
End Function
Public Function PKEY_Sync_ItemName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCE50C159, &H2FB8, &H41FD, &HBE, &H68, &HD3, &HE0, &H42, &HE2, &H74, &HBC, 3)
PKEY_Sync_ItemName = pkk
End Function
Public Function PKEY_Sync_ProgressPercentage() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 23)
PKEY_Sync_ProgressPercentage = pkk
End Function
Public Function PKEY_Sync_State() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 24)
PKEY_Sync_State = pkk
End Function
Public Function PKEY_Sync_Status() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H7BD5533E, &HAF15, &H44DB, &HB8, &HC8, &HBD, &H66, &H24, &HE1, &HD0, &H32, 10)
PKEY_Sync_Status = pkk
End Function
Public Function PKEY_Task_BillingInformation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HD37D52C6, &H261C, &H4303, &H82, &HB3, &H8, &HB9, &H26, &HAC, &H6F, &H12, 100)
PKEY_Task_BillingInformation = pkk
End Function
Public Function PKEY_Task_CompletionStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H84D8A0A, &HE6D5, &H40DE, &HBF, &H1F, &HC8, &H82, &HE, &H7C, &H87, &H7C, 100)
PKEY_Task_CompletionStatus = pkk
End Function
Public Function PKEY_Task_Owner() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8C7CC5F, &H60F2, &H4494, &HAD, &H75, &H55, &HE3, &HE0, &HB5, &HAD, &HD0, 100)
PKEY_Task_Owner = pkk
End Function
Public Function PKEY_Video_Compression() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 10)
PKEY_Video_Compression = pkk
End Function
Public Function PKEY_Video_Director() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440492, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 20)
PKEY_Video_Director = pkk
End Function
Public Function PKEY_Video_EncodingBitrate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 8)
PKEY_Video_EncodingBitrate = pkk
End Function
Public Function PKEY_Video_FourCC() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 44)
PKEY_Video_FourCC = pkk
End Function
Public Function PKEY_Video_FrameHeight() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 4)
PKEY_Video_FrameHeight = pkk
End Function
Public Function PKEY_Video_FrameRate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 6)
PKEY_Video_FrameRate = pkk
End Function
Public Function PKEY_Video_FrameWidth() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 3)
PKEY_Video_FrameWidth = pkk
End Function
Public Function PKEY_Video_HorizontalAspectRatio() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 42)
PKEY_Video_HorizontalAspectRatio = pkk
End Function
Public Function PKEY_Video_IsStereo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 98)
PKEY_Video_IsStereo = pkk
End Function
Public Function PKEY_Video_Orientation() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 99)
PKEY_Video_Orientation = pkk
End Function
Public Function PKEY_Video_SampleSize() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 9)
PKEY_Video_SampleSize = pkk
End Function
Public Function PKEY_Video_StreamName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 2)
PKEY_Video_StreamName = pkk
End Function
Public Function PKEY_Video_StreamNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 11)
PKEY_Video_StreamNumber = pkk
End Function
Public Function PKEY_Video_TotalBitrate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 43)
PKEY_Video_TotalBitrate = pkk
End Function
Public Function PKEY_Video_TranscodedForSync() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 46)
PKEY_Video_TranscodedForSync = pkk
End Function
Public Function PKEY_Video_VerticalAspectRatio() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H64440491, &H4C8B, &H11D1, &H8B, &H70, &H8, &H0, &H36, &HB1, &H1A, &H3, 45)
PKEY_Video_VerticalAspectRatio = pkk
End Function
Public Function PKEY_Volume_FileSystem() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 4)
PKEY_Volume_FileSystem = pkk
End Function
Public Function PKEY_Volume_IsMappedDrive() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H149C0B69, &H2C2D, &H48FC, &H80, &H8F, &HD3, &H18, &HD7, &H8C, &H46, &H36, 2)
PKEY_Volume_IsMappedDrive = pkk
End Function
Public Function PKEY_Volume_IsRoot() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H9B174B35, &H40FF, &H11D2, &HA2, &H7E, &H0, &HC0, &H4F, &HC3, &H8, &H71, 10)
PKEY_Volume_IsRoot = pkk
End Function

'=============================================================================================
'
'MISSING PROPERYKEYS
'
'The following PROPERTYKEY entries are enumerated by IPropertySystem and related, and present
'in Explorer and some are commonly used, but have been omitted from propkey.h in both the 7.1
'release and the Windows 10 SDK release. These are not 3rd party either.
'=============================================================================================
Public Function PKEY_Software_ProductVersion() As PROPERTYKEY
'{0CEF7D53-FA64-11D1-A203-0000F81FEDEE, 8
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &HCEF7D53, CInt(&HFA64), CInt(&H11D1), &HA2, &H3, &H0, &H0, &HF8, &H1F, &HED, &HEE, 8)
 PKEY_Software_ProductVersion = iid
End Function
Public Function PKEY_Software_SupportURL() As PROPERTYKEY
'{841E4F90-FF59-4D16-8947-E81BBFFAB36D, 6
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H841E4F90, CInt(&HFF59), CInt(&H4D16), &H89, &H47, &HE8, &H1B, &HBF, &HFA, &HB3, &H6D, 6)
 PKEY_Software_SupportURL = iid
End Function
Public Function PKEY_Contact_GivenName() As PROPERTYKEY
'{176DC63C-2688-4E89-8143-A347800F25E9, 70
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H176DC63C, CInt(&H2688), CInt(&H4E89), &H81, &H43, &HA3, &H47, &H80, &HF, &H25, &HE9, 70)
 PKEY_Contact_GivenName = iid
End Function
Public Function PKEY_ItemSearchLocation() As PROPERTYKEY
'{23620678-CCD4-47C0-9963-95A8405678A3, 100
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H23620678, CInt(&HCCD4), CInt(&H47C0), &H99, &H63, &H95, &HA8, &H40, &H56, &H78, &HA3, 100)
 PKEY_ItemSearchLocation = iid
End Function
'The following may only be installed because of XP Mode or MS VirtualPC, I'm including them just in case they're standard
Public Function PKEY_Microsoft_VirtualMachine_PrimaryVHD() As PROPERTYKEY
'{DAB567AE-62BE-4188-B5F2-B10ADF3E2AF2, 100
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &HDAB567AE, CInt(&H62BE), CInt(&H4188), &HB5, &HF2, &HB1, &HA, &HDF, &H3E, &H2A, &HF2, 100)
 PKEY_Microsoft_VirtualMachine_PrimaryVHD = iid
End Function
Public Function PKEY_Microsoft_VirtualMachine_RAM() As PROPERTYKEY
'{B3414030-57A1-453A-A069-F0024026C58C, 101
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &HB3414030, CInt(&H57A1), CInt(&H453A), &HA0, &H69, &HF0, &H2, &H40, &H26, &HC5, &H8C, 101)
 PKEY_Microsoft_VirtualMachine_RAM = iid
End Function
Public Function PKEY_Microsoft_VirtualMachine_Status() As PROPERTYKEY
'{6B585045-57FF-43A0-8731-E2DDF3F5D6EC, 103
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H6B585045, CInt(&H57FF), CInt(&H43A0), &H87, &H31, &HE2, &HDD, &HF3, &HF5, &HD6, &HEC, 103)
 PKEY_Microsoft_VirtualMachine_Status = iid
End Function
Public Function PKEY_Microsoft_VirtualMachine_VMPath() As PROPERTYKEY
'{98E8009B-E40E-4820-BEDF-7AAB22FD9BED, 102
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H98E8009B, CInt(&HE40E), CInt(&H4820), &HBE, &HDF, &H7A, &HAB, &H22, &HFD, &H9B, &HED, 102)
 PKEY_Microsoft_VirtualMachine_VMPath = iid
End Function


'========================================================================
'KEYS FROM OTHER FILES

Public Function PKEY_NAME() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HB725F130, &H47EF, &H101A, &HA5, &HF1, &H2, &H60, &H8C, &H9E, &HEB, &HAC, 10)
 'DEVPROP_TYPE_STRING
PKEY_NAME = pkk
End Function
Public Function PKEY_Device_DeviceDesc() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 2)
 'DEVPROP_TYPE_STRING
PKEY_Device_DeviceDesc = pkk
End Function
Public Function PKEY_Device_HardwareIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 3)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_HardwareIds = pkk
End Function
Public Function PKEY_Device_CompatibleIds() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 4)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_CompatibleIds = pkk
End Function
Public Function PKEY_Device_Service() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 6)
 'DEVPROP_TYPE_STRING
PKEY_Device_Service = pkk
End Function
Public Function PKEY_Device_Class() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 9)
 'DEVPROP_TYPE_STRING
PKEY_Device_Class = pkk
End Function
Public Function PKEY_Device_ClassGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 10)
 'DEVPROP_TYPE_GUID
PKEY_Device_ClassGuid = pkk
End Function
Public Function PKEY_Device_Driver() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 11)
 'DEVPROP_TYPE_STRING
PKEY_Device_Driver = pkk
End Function
Public Function PKEY_Device_ConfigFlags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 12)
 'DEVPROP_TYPE_UINT32
PKEY_Device_ConfigFlags = pkk
End Function
Public Function PKEY_Device_Manufacturer() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 13)
 'DEVPROP_TYPE_STRING
PKEY_Device_Manufacturer = pkk
End Function
Public Function PKEY_Device_FriendlyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 14)
 'DEVPROP_TYPE_STRING
PKEY_Device_FriendlyName = pkk
End Function
Public Function PKEY_Device_LocationInfo() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 15)
 'DEVPROP_TYPE_STRING
PKEY_Device_LocationInfo = pkk
End Function
Public Function PKEY_Device_PDOName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 16)
 'DEVPROP_TYPE_STRING
PKEY_Device_PDOName = pkk
End Function
Public Function PKEY_Device_Capabilities() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 17)
 'DEVPROP_TYPE_UNINT32
PKEY_Device_Capabilities = pkk
End Function
Public Function PKEY_Device_UINumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 18)
 'DEVPROP_TYPE_STRING
PKEY_Device_UINumber = pkk
End Function
Public Function PKEY_Device_UpperFilters() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 19)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_UpperFilters = pkk
End Function
Public Function PKEY_Device_LowerFilters() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 20)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_LowerFilters = pkk
End Function
Public Function PKEY_Device_BusTypeGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 21)
 'DEVPROP_TYPE_GUID
PKEY_Device_BusTypeGuid = pkk
End Function
Public Function PKEY_Device_LegacyBusType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 22)
 'DEVPROP_TYPE_UINT32
PKEY_Device_LegacyBusType = pkk
End Function
Public Function PKEY_Device_BusNumber() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 23)
 'DEVPROP_TYPE_UINT32
PKEY_Device_BusNumber = pkk
End Function
Public Function PKEY_Device_EnumeratorName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 24)
 'DEVPROP_TYPE_STRING
PKEY_Device_EnumeratorName = pkk
End Function
Public Function PKEY_Device_Security() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 25)
 'DEVPROP_TYPE_SECURITY_DESCRIPTOR
PKEY_Device_Security = pkk
End Function
Public Function PKEY_Device_SecuritySDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 26)
 'DEVPROP_TYPE_SECURITY_DESCRIPTOR_STRING
PKEY_Device_SecuritySDS = pkk
End Function
Public Function PKEY_Device_DevType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 27)
 'DEVPROP_TYPE_UINT32
PKEY_Device_DevType = pkk
End Function
Public Function PKEY_Device_Exclusive() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 28)
 'DEVPROP_TYPE_UINT32
PKEY_Device_Exclusive = pkk
End Function
Public Function PKEY_Device_Characteristics() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 29)
 'DEVPROP_TYPE_UINT32
PKEY_Device_Characteristics = pkk
End Function
Public Function PKEY_Device_Address() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 30)
 'DEVPROP_TYPE_UINT32
PKEY_Device_Address = pkk
End Function
Public Function PKEY_Device_UINumberDescFormat() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 31)
 'DEVPROP_TYPE_STRING
PKEY_Device_UINumberDescFormat = pkk
End Function
Public Function PKEY_Device_PowerData() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 32)
 'DEVPROP_TYPE_BINARY
PKEY_Device_PowerData = pkk
End Function
Public Function PKEY_Device_RemovalPolicy() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 33)
 'DEVPROP_TYPE_UINT32
PKEY_Device_RemovalPolicy = pkk
End Function
Public Function PKEY_Device_RemovalPolicyDefault() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 34)
 'DEVPROP_TYPE_UINT32
PKEY_Device_RemovalPolicyDefault = pkk
End Function
Public Function PKEY_Device_RemovalPolicyOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 35)
 'DEVPROP_TYPE_UINT32
PKEY_Device_RemovalPolicyOverride = pkk
End Function
Public Function PKEY_Device_InstallState() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 36)
 'DEVPROP_TYPE_UINT32
PKEY_Device_InstallState = pkk
End Function
Public Function PKEY_Device_LocationPaths() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 37)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_LocationPaths = pkk
End Function
Public Function PKEY_Device_BaseContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA45C254E, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0, 38)
 'DEVPROP_TYPE_GUID
PKEY_Device_BaseContainerId = pkk
End Function
Public Function PKEY_Device_DevNodeStatus() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 2)
 'DEVPROP_TYPE_UINT32
PKEY_Device_DevNodeStatus = pkk
End Function
Public Function PKEY_Device_ProblemCode() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 3)
 'DEVPROP_TYPE_UINT32
PKEY_Device_ProblemCode = pkk
End Function
Public Function PKEY_Device_EjectionRelations() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 4)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_EjectionRelations = pkk
End Function
Public Function PKEY_Device_RemovalRelations() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 5)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_RemovalRelations = pkk
End Function
Public Function PKEY_Device_PowerRelations() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 6)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_PowerRelations = pkk
End Function
Public Function PKEY_Device_BusRelations() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 7)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_BusRelations = pkk
End Function
Public Function PKEY_Device_Parent() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 8)
 'DEVPROP_TYPE_STRING
PKEY_Device_Parent = pkk
End Function
Public Function PKEY_Device_Children() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 9)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_Children = pkk
End Function
Public Function PKEY_Device_Siblings() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 10)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_Siblings = pkk
End Function
Public Function PKEY_Device_TransportRelations() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4340A6C5, &H93FA, &H4706, &H97, &H2C, &H7B, &H64, &H80, &H8, &HA5, &HA7, 11)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_TransportRelations = pkk
End Function
Public Function PKEY_Device_Reported() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80497100, &H8C73, &H48B9, &HAA, &HD9, &HCE, &H38, &H7E, &H19, &HC5, &H6E, 2)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_Reported = pkk
End Function
Public Function PKEY_Device_Legacy() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80497100, &H8C73, &H48B9, &HAA, &HD9, &HCE, &H38, &H7E, &H19, &HC5, &H6E, 3)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_Legacy = pkk
End Function
Public Function PKEY_Device_InstanceId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H78C34FC8, &H104A, &H4ACA, &H9E, &HA4, &H52, &H4D, &H52, &H99, &H6E, &H57, 256)
 'DEVPROP_TYPE_STRING
PKEY_Device_InstanceId = pkk
End Function
Public Function PKEY_Device_ContainerId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H8C7ED206, &H3F8A, &H4827, &HB3, &HAB, &HAE, &H9E, &H1F, &HAE, &HFC, &H6C, 2)
 'DEVPROP_TYPE_GUID
PKEY_Device_ContainerId = pkk
End Function
Public Function PKEY_Device_ModelId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 2)
 'DEVPROP_TYPE_GUID
PKEY_Device_ModelId = pkk
End Function
Public Function PKEY_Device_FriendlyNameAttributes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 3)
 'DEVPROP_TYPE_UINT32
PKEY_Device_FriendlyNameAttributes = pkk
End Function
Public Function PKEY_Device_ManufacturerAttributes() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 4)
 'DEVPROP_TYPE_UINT32
PKEY_Device_ManufacturerAttributes = pkk
End Function
Public Function PKEY_Device_PresenceNotForDevice() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 5)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_PresenceNotForDevice = pkk
End Function
Public Function PKEY_Device_SignalStrength() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 6)
 'DEVPROP_TYPE_UINT32
PKEY_Device_SignalStrength = pkk
End Function
Public Function PKEY_Device_IsAssociateableByUserAction() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H80D81EA6, &H7473, &H4B0C, &H82, &H16, &HEF, &HC1, &H1A, &H2C, &H4C, &H8B, 7)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_IsAssociateableByUserAction = pkk
End Function
Public Function PKEY_Numa_Proximity_Domain() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 1)
 'DEVPROP_TYPE_UINT32
PKEY_Numa_Proximity_Domain = pkk
End Function
Public Function PKEY_Device_DHP_Rebalance_Policy() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 2)
 'DEVPROP_TYPE_UINT32
PKEY_Device_DHP_Rebalance_Policy = pkk
End Function
Public Function PKEY_Device_Numa_Node() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 3)
 'DEVPROP_TYPE_UINT32
PKEY_Device_Numa_Node = pkk
End Function
Public Function PKEY_Device_BusReportedDeviceDesc() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H540B947E, &H8B40, &H45BC, &HA8, &HA2, &H6A, &HB, &H89, &H4C, &HBD, &HA2, 4)
 'DEVPROP_TYPE_STRING
PKEY_Device_BusReportedDeviceDesc = pkk
End Function
Public Function PKEY_Device_InstallInProgress() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H83DA6326, &H97A6, &H4088, &H94, &H53, &HA1, &H92, &H3F, &H57, &H3B, &H29, 9)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_InstallInProgress = pkk
End Function
Public Function PKEY_Device_DriverDate() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 2)
 'DEVPROP_TYPE_FILETIME
PKEY_Device_DriverDate = pkk
End Function
Public Function PKEY_Device_DriverVersion() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 3)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverVersion = pkk
End Function
Public Function PKEY_Device_DriverDesc() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 4)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverDesc = pkk
End Function
Public Function PKEY_Device_DriverInfPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 5)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverInfPath = pkk
End Function
Public Function PKEY_Device_DriverInfSection() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 6)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverInfSection = pkk
End Function
Public Function PKEY_Device_DriverInfSectionExt() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 7)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverInfSectionExt = pkk
End Function
Public Function PKEY_Device_MatchingDeviceId() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 8)
 'DEVPROP_TYPE_STRING
PKEY_Device_MatchingDeviceId = pkk
End Function
Public Function PKEY_Device_DriverProvider() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 9)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverProvider = pkk
End Function
Public Function PKEY_Device_DriverPropPageProvider() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 10)
 'DEVPROP_TYPE_STRING
PKEY_Device_DriverPropPageProvider = pkk
End Function
Public Function PKEY_Device_DriverCoInstallers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 11)
 'DEVPROP_TYPE_STRING_LIST
PKEY_Device_DriverCoInstallers = pkk
End Function
Public Function PKEY_Device_ResourcePickerTags() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 12)
 'DEVPROP_TYPE_STRING
PKEY_Device_ResourcePickerTags = pkk
End Function
Public Function PKEY_Device_ResourcePickerExceptions() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 13)
 'DEVPROP_TYPE_STRING
PKEY_Device_ResourcePickerExceptions = pkk
End Function
Public Function PKEY_Device_DriverRank() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 14)
 'DEVPROP_TYPE_UINT32
PKEY_Device_DriverRank = pkk
End Function
Public Function PKEY_Device_DriverLogoLevel() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 15)
 'DEVPROP_TYPE_UINT32
PKEY_Device_DriverLogoLevel = pkk
End Function
Public Function PKEY_Device_NoConnectSound() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 17)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_NoConnectSound = pkk
End Function
Public Function PKEY_Device_GenericDriverInstalled() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 18)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_GenericDriverInstalled = pkk
End Function
Public Function PKEY_Device_AdditionalSoftwareRequested() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HA8B865DD, &H2E3D, &H4094, &HAD, &H97, &HE5, &H93, &HA7, &HC, &H75, &HD6, 19)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_AdditionalSoftwareRequested = pkk
End Function
Public Function PKEY_Device_SafeRemovalRequired() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFD97640, &H86A3, &H4210, &HB6, &H7C, &H28, &H9C, &H41, &HAA, &HBE, &H55, 2)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_SafeRemovalRequired = pkk
End Function
Public Function PKEY_Device_SafeRemovalRequiredOverride() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HAFD97640, &H86A3, &H4210, &HB6, &H7C, &H28, &H9C, &H41, &HAA, &HBE, &H55, 3)
 'DEVPROP_TYPE_BOOLEAN
PKEY_Device_SafeRemovalRequiredOverride = pkk
End Function
Public Function PKEY_DrvPkg_Model() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 2)
 'DEVPROP_TYPE_STRING
PKEY_DrvPkg_Model = pkk
End Function
Public Function PKEY_DrvPkg_VendorWebSite() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 3)
 'DEVPROP_TYPE_STRING
PKEY_DrvPkg_VendorWebSite = pkk
End Function
Public Function PKEY_DrvPkg_DetailedDescription() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 4)
 'DEVPROP_TYPE_STRING
PKEY_DrvPkg_DetailedDescription = pkk
End Function
Public Function PKEY_DrvPkg_DocumentationLink() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 5)
 'DEVPROP_TYPE_STRING
PKEY_DrvPkg_DocumentationLink = pkk
End Function
Public Function PKEY_DrvPkg_Icon() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 6)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DrvPkg_Icon = pkk
End Function
Public Function PKEY_DrvPkg_BrandingIcon() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &HCF73BB51, &H3ABF, &H44A2, &H85, &HE0, &H9A, &H3D, &HC7, &HA1, &H21, &H32, 7)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DrvPkg_BrandingIcon = pkk
End Function
Public Function PKEY_DeviceClass_UpperFilters() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 19)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DeviceClass_UpperFilters = pkk
End Function
Public Function PKEY_DeviceClass_LowerFilters() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 20)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DeviceClass_LowerFilters = pkk
End Function
Public Function PKEY_DeviceClass_Security() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 25)
 'DEVPROP_TYPE_SECURITY_DESCRIPTOR
PKEY_DeviceClass_Security = pkk
End Function
Public Function PKEY_DeviceClass_SecuritySDS() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 26)
 'DEVPROP_TYPE_SECURITY_DESCRIPTOR_STRING
PKEY_DeviceClass_SecuritySDS = pkk
End Function
Public Function PKEY_DeviceClass_DevType() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 27)
 'DEVPROP_TYPE_UINT32
PKEY_DeviceClass_DevType = pkk
End Function
Public Function PKEY_DeviceClass_Exclusive() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 28)
 'DEVPROP_TYPE_UINT32
PKEY_DeviceClass_Exclusive = pkk
End Function
Public Function PKEY_DeviceClass_Characteristics() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H4321918B, &HF69E, &H470D, &HA5, &HDE, &H4D, &H88, &HC7, &H5A, &HD2, &H4B, 29)
 'DEVPROP_TYPE_UINT32
PKEY_DeviceClass_Characteristics = pkk
End Function
Public Function PKEY_DeviceClass_Name() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 2)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_Name = pkk
End Function
Public Function PKEY_DeviceClass_ClassName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 3)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_ClassName = pkk
End Function
Public Function PKEY_DeviceClass_Icon() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 4)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_Icon = pkk
End Function
Public Function PKEY_DeviceClass_ClassInstaller() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 5)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_ClassInstaller = pkk
End Function
Public Function PKEY_DeviceClass_PropPageProvider() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 6)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_PropPageProvider = pkk
End Function
Public Function PKEY_DeviceClass_NoInstallClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 7)
 'DEVPROP_TYPE_BOOLEAN
PKEY_DeviceClass_NoInstallClass = pkk
End Function
Public Function PKEY_DeviceClass_NoDisplayClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 8)
 'DEVPROP_TYPE_BOOLEAN
PKEY_DeviceClass_NoDisplayClass = pkk
End Function
Public Function PKEY_DeviceClass_SilentInstall() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 9)
 'DEVPROP_TYPE_BOOLEAN
PKEY_DeviceClass_SilentInstall = pkk
End Function
Public Function PKEY_DeviceClass_NoUseClass() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 10)
 'DEVPROP_TYPE_BOOLEAN
PKEY_DeviceClass_NoUseClass = pkk
End Function
Public Function PKEY_DeviceClass_DefaultService() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 11)
 'DEVPROP_TYPE_STRING
PKEY_DeviceClass_DefaultService = pkk
End Function
Public Function PKEY_DeviceClass_IconPath() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H259ABFFC, &H50A7, &H47CE, &HAF, &H8, &H68, &HC9, &HA7, &HD7, &H33, &H66, 12)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DeviceClass_IconPath = pkk
End Function
Public Function PKEY_DeviceClass_ClassCoInstallers() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H713D1703, &HA2E2, &H49F5, &H92, &H14, &H56, &H47, &H2E, &HF3, &HDA, &H5C, 2)
 'DEVPROP_TYPE_STRING_LIST
PKEY_DeviceClass_ClassCoInstallers = pkk
End Function
Public Function PKEY_DeviceInterface_FriendlyName() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 2)
 'DEVPROP_TYPE_STRING
PKEY_DeviceInterface_FriendlyName = pkk
End Function
Public Function PKEY_DeviceInterface_Enabled() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 3)
 'DEVPROP_TYPE_BOOLEAN
PKEY_DeviceInterface_Enabled = pkk
End Function
Public Function PKEY_DeviceInterface_ClassGuid() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22, 4)
 'DEVPROP_TYPE_GUID
PKEY_DeviceInterface_ClassGuid = pkk
End Function
Public Function PKEY_DeviceInterfaceClass_DefaultInterface() As PROPERTYKEY
Static pkk As PROPERTYKEY
 If (pkk.fmtid.Data1 = 0&) Then Call DEFINE_PROPERTYKEY(pkk, &H14C83A99, &HB3F, &H44B7, &HBE, &H4C, &HA1, &H78, &HD3, &H99, &H5, &H64, 2)
 'DEVPROP_TYPE_STRING
PKEY_DeviceInterfaceClass_DefaultInterface = pkk
End Function
'========================================================================
'UNDOCUMENTED PROPERTY KEYS
'The most mysterious of the most mysterious

Public Function PKEY_Software_DateInstalled() As PROPERTYKEY
'{841E4F90-FF59-4D16-8947-E81BBFFAB36D},11
Static iid As PROPERTYKEY
 If (iid.fmtid.Data1 = 0) Then Call DEFINE_PROPERTYKEY(iid, &H841E4F90, CInt(&HFF59), CInt(&H4D16), &H89, &H47, &HE8, &H1B, &HBF, &HFA, &HB3, &H6D, 11)
 PKEY_Software_DateInstalled = iid

End Function
