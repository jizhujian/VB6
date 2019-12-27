Attribute VB_Name = "mDirectShow"
Option Explicit

Public Function LIBID_QuartzNetTypeLib() As UUID
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B1, &HAD4, &H11CE, &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
LIBID_QuartzNetTypeLib = iid
End Function

Public Function IID_IAMCollection() As UUID
'{56A868B9-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B9, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IAMCollection = iid
End Function
Public Function IID_IMediaControl() As UUID
'{56A868B1-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B1, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaControl = iid
End Function
Public Function IID_IMediaEvent() As UUID
'{56A868B6-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B6, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaEvent = iid
End Function
Public Function IID_IMediaEventEx() As UUID
'{56A868C0-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868C0, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaEventEx = iid
End Function
Public Function IID_IMediaPosition() As UUID
'{56A868B2-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B2, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaPosition = iid
End Function
Public Function IID_IBasicAudio() As UUID
'{56A868B3-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B3, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBasicAudio = iid
End Function
Public Function IID_IVideoWindow() As UUID
'{56A868B4-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B4, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IVideoWindow = iid
End Function
Public Function IID_IBasicVideo() As UUID
'{56A868B5-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B5, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBasicVideo = iid
End Function
Public Function IID_IBasicVideo2() As UUID
'{329BB360-F6EA-11D1-9038-00A0C9697298}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H329BB360, CInt(&HF6EA), CInt(&H11D1), &H90, &H38, &H0, &HA0, &HC9, &H69, &H72, &H98)
IID_IBasicVideo2 = iid
End Function
Public Function IID_IDeferredCommand() As UUID
'{56A868B8-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B8, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IDeferredCommand = iid
End Function
Public Function IID_IQueueCommand() As UUID
'{56A868B7-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868B7, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IQueueCommand = iid
End Function
Public Function IID_IFilterInfo() As UUID
'{E436EBB3-524F-11CE-9F53-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HE436EBB3, CInt(&H524F), CInt(&H11CE), &H9F, &H53, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IFilterInfo = iid
End Function
Public Function IID_IRegFilterInfo() As UUID
'{56A868BB-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BB, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IRegFilterInfo = iid
End Function
Public Function IID_IMediaTypeInfo() As UUID
'{56A868BC-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BC, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaTypeInfo = iid
End Function
Public Function IID_IPinInfo() As UUID
'{56A868BD-0AD4-11CE-B03A-0020AF0BA770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868BD, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IPinInfo = iid
End Function
Public Function IID_IAMStats() As UUID
'{BC9BCF80-DCD2-11D2-ABF6-00A0C905F375}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBC9BCF80, CInt(&HDCD2), CInt(&H11D2), &HAB, &HF6, &H0, &HA0, &HC9, &H5, &HF3, &H75)
IID_IAMStats = iid
End Function
Public Function IID_IEnumMediaTypes() As UUID
'{89c31040-846b-11ce-97d3-00aa0055595a}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H89C31040, CInt(&H846B), CInt(&H11CE), &H97, &HD3, &H0, &HAA, &H0, &H55, &H59, &H5A)
IID_IEnumMediaTypes = iid
End Function
Public Function IID_IPin() As UUID
'{56a86891-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86891, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IPin = iid
End Function
Public Function IID_IEnumPins() As UUID
'{56a86892-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86892, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IEnumPins = iid
End Function
Public Function IID_IReferenceClock() As UUID
'{56a86897-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86897, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IReferenceClock = iid
End Function
Public Function IID_IMediaFilter() As UUID
'{56a86899-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86899, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IMediaFilter = iid
End Function
Public Function IID_IBaseFilter() As UUID
'{56a86895-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86895, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IBaseFilter = iid
End Function
Public Function IID_IEnumFilters() As UUID
'{56a86893-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A86893, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IEnumFilters = iid
End Function
Public Function IID_IFilterGraph() As UUID
'{56a8689f-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A8689F, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IFilterGraph = iid
End Function
Public Function IID_IFileSinkFilter() As UUID
'{a2104830-7c70-11cf-8bce-00aa00a3f1a6}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HA2104830, CInt(&H7C70), CInt(&H11CF), &H8B, &HCE, &H0, &HAA, &H0, &HA3, &HF1, &HA6)
IID_IFileSinkFilter = iid
End Function
Public Function IID_IAMCopyCaptureFileProgress() As UUID
'{670d1d20-a068-11d0-b3f0-00aa003761c5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H670D1D20, CInt(&HA068), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMCopyCaptureFileProgress = iid
End Function
Public Function IID_IGraphBuilder() As UUID
'{56a868a9-0ad4-11ce-b03a-0020af0ba770}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56A868A9, CInt(&HAD4), CInt(&H11CE), &HB0, &H3A, &H0, &H20, &HAF, &HB, &HA7, &H70)
IID_IGraphBuilder = iid
End Function
Public Function IID_ICaptureGraphBuilder() As UUID
'{bf87b6e0-8c27-11d0-b3f0-00aa003761c5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HBF87B6E0, CInt(&H8C27), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_ICaptureGraphBuilder = iid
End Function
Public Function IID_ICaptureGraphBuilder2() As UUID
'{93E5A4E0-2D50-11d2-ABFA-00A0C9C6E38D}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H93E5A4E0, CInt(&H2D50), CInt(&H11D2), &HAB, &HFA, &H0, &HA0, &HC9, &HC6, &HE3, &H8D)
IID_ICaptureGraphBuilder2 = iid
End Function
Public Function IID_IAMChannelInfo() As UUID
'{FA2AA8F1-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F1, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMChannelInfo = iid
End Function
Public Function IID_IAMNetworkStatus() As UUID
'{FA2AA8F3-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F3, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMNetworkStatus = iid
End Function
Public Function IID_IAMNetShowExProps() As UUID
'{FA2AA8F5-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F5, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMNetShowExProps = iid
End Function
Public Function IID_IAMExtendedErrorInfo() As UUID
'{FA2AA8F6-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F6, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMExtendedErrorInfo = iid
End Function
Public Function IID_IAMNetShowPreroll() As UUID
'{AAE7E4E2-6388-11D1-8D93-006097C9A2B2}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HAAE7E4E2, CInt(&H6388), CInt(&H11D1), &H8D, &H93, &H0, &H60, &H97, &HC9, &HA2, &HB2)
IID_IAMNetShowPreroll = iid
End Function
Public Function IID_IAMMediaContent() As UUID
'{FA2AA8F4-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F4, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMMediaContent = iid
End Function
Public Function IID_IAMExtendedSeeking() As UUID
'{FA2AA8F9-8B62-11D0-A520-000000000000}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HFA2AA8F9, CInt(&H8B62), CInt(&H11D0), &HA5, &H20, &H0, &H0, &H0, &H0, &H0, &H0)
IID_IAMExtendedSeeking = iid
End Function
Public Function IID_IAMMediaContent2() As UUID
'{CE8F78C1-74D9-11D2-B09D-00A0C9A81117}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCE8F78C1, CInt(&H74D9), CInt(&H11D2), &HB0, &H9D, &H0, &HA0, &HC9, &HA8, &H11, &H17)
IID_IAMMediaContent2 = iid
End Function
Public Function IID_IAMAnalogVideoDecoder() As UUID
'{C6E13350-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13350, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMAnalogVideoDecoder = iid
End Function
Public Function IID_IAMAsyncReaderTimestampScaling() As UUID
'{cf7b26fc-9a00-485b-8147-3e789d5e8f67}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HCF7B26FC, CInt(&H9A00), CInt(&H485B), &H81, &H47, &H3E, &H78, &H9D, &H5E, &H8F, &H67)
IID_IAMAsyncReaderTimestampScaling = iid
End Function
Public Function IID_IAMAudioInputMixer() As UUID
'{54C39221-8380-11d0-B3F0-00AA003761C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H54C39221, CInt(&H8380), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMAudioInputMixer = iid
End Function
Public Function IID_IAMAudioRendererStats() As UUID
'{22320CB2-D41A-11d2-BF7C-D7CB9DF0BF93}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H22320CB2, CInt(&HD41A), CInt(&H11D2), &HBF, &H7C, &HD7, &HCB, &H9D, &HF0, &HBF, &H93)
IID_IAMAudioRendererStats = iid
End Function
Public Function IID_IAMBufferNegotiation() As UUID
'{56ED71A0-AF5F-11D0-B3F0-00AA003761C5}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H56ED71A0, CInt(&HAF5F), CInt(&H11D0), &HB3, &HF0, &H0, &HAA, &H0, &H37, &H61, &HC5)
IID_IAMBufferNegotiation = iid
End Function
Public Function IID_IAMCameraControl() As UUID
'{C6E13370-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13370, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMCameraControl = iid
End Function
Public Function IID_IAMCertifiedOutputProtection() As UUID
'{6feded3e-0ff1-4901-a2f1-43f7012c8515}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H6FEDED3E, CInt(&HFF1), CInt(&H4901), &HA2, &HF1, &H43, &HF7, &H1, &H2C, &H85, &H15)
IID_IAMCertifiedOutputProtection = iid
End Function
Public Function IID_IAMClockAdjust() As UUID
'{4d5466b0-a49c-11d1-abe8-00a0c905f375}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H4D5466B0, CInt(&HA49C), CInt(&H11D1), &HAB, &HE8, &H0, &HA0, &HC9, &H5, &HF3, &H75)
IID_IAMClockAdjust = iid
End Function
Public Function IID_IAMClockSlave() As UUID
'{9FD52741-176D-4b36-8F51-CA8F933223BE}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &H9FD52741, CInt(&H176D), CInt(&H4B36), &H8F, &H51, &HCA, &H8F, &H93, &H32, &H23, &HBE)
IID_IAMClockSlave = iid
End Function
Public Function IID_IAMCrossbar() As UUID
'{C6E13380-30AC-11d0-A18C-00A0C9118956}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC6E13380, CInt(&H30AC), CInt(&H11D0), &HA1, &H8C, &H0, &HA0, &HC9, &H11, &H89, &H56)
IID_IAMCrossbar = iid
End Function
Public Function IID_IAMDecoderCaps() As UUID
'{c0dff467-d499-4986-972b-e1d9090fa941}
Static iid As UUID
 If (iid.Data1 = 0&) Then Call DEFINE_UUID(iid, &HC0DFF467, CInt(&HD499), CInt(&H4986), &H97, &H2B, &HE1, &HD9, &H9, &HF, &HA9, &H41)
IID_IAMDecoderCaps = iid
End Function
