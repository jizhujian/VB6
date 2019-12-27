''' <summary>
''' 农历函数库
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(Lunar.ClassId, Lunar.InterfaceId, Lunar.EventsId)> _
Public Class Lunar

  ''' <summary>
  ''' COM注册必须
  ''' </summary>
  ''' <remarks></remarks>
  Public Const ClassId As String = "49f6c80f-8688-482a-bf16-56eb7f18099e"
  Public Const InterfaceId As String = "cf625e6f-6bac-4ced-ab0a-50d03fc302ae"
  Public Const EventsId As String = "1922f340-5375-4ce8-bfce-f24ece736b18"

  Public Sub New()
    MyBase.New()
  End Sub

#Region "变量声明"

  Private _solarDate As Date
  Private _lunarYear As Integer
  Private _lunarMonth As Integer
  Private _lunarDay As Integer
  Private _isLeapMonth As Boolean
  Private _heavenlyStem As Integer
  Private _earthlyBranch As Integer
  Private _leepMonth As Integer
  Private _solarTerm As Integer
  Private _solarTermDate As Date

  Private Const LUNAR_BEGINNINGYEAR As Integer = -849
  Private Const LUNAR_LEAPMONTH As String = "0c0080050010a0070030c0080050010a0070030c0080050020a" & _
    "0070030c0080050020a0070030c0090050020a0070030c0090050020a0060030c0060030c00900600c0c00" & _
    "60c00c00c00c0c000600c0c0006090303030006000c00c060c0006c00000c0c0c0060003030006c00009009" & _
    "c0090c00c009000300030906030030c0c00060c00090c0060600c0030060c00c003006009060030c006006" & _
    "0c0090900c00090c0090c00c006030006060003030c0c00030c0060030c0090060030c0090300c00800500" & _
    "20a0060030c0080050020b0070030c0090050010a0070030b0090060020a0070040c0080050020a006003" & _
    "0c0080050020b0070030c0090050010a0070030b0090060020a0070040c0080050020a0060030c0080050" & _
    "020b0070030c0090050000c00900909009009090090090090900900909009009009090090090900900900" & _
    "909009009090090090090900900909009009090090090090900900909009009009090090090900900900" & _
    "909009009090060030c0090050010a0070030b008005001090070040c0080050020a0060030c009004001" & _
    "0a0060030c0090050010a0070030b0080050010a008005001090050020a0060030c0080040010a0060030" & _
    "c0090050010a0070030b0080050010a0070030b008005001090070040c0080050020a0060030c00800400" & _
    "10a0060030c0090050010a0070030b008005001090070040c0080050020a0060030c0080040010a006003" & _
    "0c0090050010a0060030c0090050010a0070030b008005001090070040c0080050020a0060030c0080040" & _
    "010a0070030b0080050010a0070040c0080050020a0060030c0080040010a0070030c0090050010a00700" & _
    "30b0080050020a0060030c0080040010a0060030c0090050050020a0060030c0090050010b0070030c009" & _
    "0050010a0070040c0080040020a0060030c0080050020a0060030c0090050010a0070030b0080040020a0" & _
    "060040c0090050020b0070030c00a0050010a0070030b0090050020a0070030c0080040020a0060030c00" & _
    "90050010a0070030c0090050030b007005001090050020a007004001090060020c0070050c0090060030b" & _
    "0080040020a0060030b0080040010a0060030b0080050010a0050040c0080050010a0060030c008005001" & _
    "0b0070030c007005001090070030b0070040020a0060030c0080040020a0070030b0090050010a0060040" & _
    "c0080050020a0060040c0080050010b0070030c007005001090070030c0080050020a0070030c009005002" & _
    "0a0070030c0090050020a0060040c0090050020a0060040c0090050010b0070030c0080050030b00700400" & _
    "1090060020c008004002090060020a008004001090050030b0080040020a0060040b0080040c00a0060020" & _
    "b007005001090060030b0070050020a0060020c008004002090070030c008005002090070040c008004002" & _
    "0a0060040b0090050010a0060030b0080050020a0060040c0080050010b00700300108005001090070030c" & _
    "0080050020a007003001090050030a0070030b0090050020a0060040c0090050030b0070040c0090050010" & _
    "c0070040c0080060020b00700400a090060020b007003002090060020a005004001090050030b007004001" & _
    "090050040c0080040c00a0060020c007005001090060030b0070050020a0060020c008004002090060030b" & _
    "008004002090060030b0080040020a0060040b0080040010b0060030b0070050010a006004002070050030" & _
    "8006004003070050030700600400307005003080060040030700500409006004003070050040900600500" & _
    "2070050030a006005003070050040020600400206005003002060040030700500409006004003070050040" & _
    "8007005003080050040a006005003070050040020600500308005004002060050020700500400206005003" & _
    "07006004002070050030800600400307005004080060040a00600500308005004002070050040900600400" & _
    "2060050030b0060050020700500308006004003070050040800600400307005004080060040020"

#End Region

#Region "属性"

  ''' <summary>
  ''' 公历日期
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property SolarDate() As Date
    Get
      Return _solarDate
    End Get
    Set(ByVal value As Date)
      _solarDate = value
      CalculateLunar()
    End Set
  End Property

  ''' <summary>
  ''' 农历年
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarYear() As Integer
    Get
      Return _lunarYear
    End Get
  End Property

  ''' <summary>
  ''' 农历年中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarYearName() As String
    Get
      Dim lunarYearNames As String() = New String() {"○", "一", "二", "三", "四", "五", "六", "七", "八", "九"}
      Dim yearName As String
      Dim i As Integer = _lunarYear
      Do While i > 0
        yearName = lunarYearNames(i Mod 10) & yearName
        i \= 10
      Loop
      Return yearName & "年"
    End Get
  End Property

  ''' <summary>
  ''' 是否闰月
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property IsLeapMonth() As Boolean
    Get
      Return _isLeapMonth
    End Get
  End Property

  ''' <summary>
  ''' 是否闰月中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property IsLeapMonthName() As String
    Get
      If _isLeapMonth Then Return "闰"
      Return String.Empty
    End Get
  End Property

  ''' <summary>
  ''' 农历月
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarMonth() As Integer
    Get
      Return _lunarMonth
    End Get
  End Property

  ''' <summary>
  ''' 农历月中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarMonthName() As String
    Get
      Dim lunarMonthNames As String() = New String() {"正", "二", "三", "四", "五", "六", "七", "八", "九", "十", "十一", "腊"}
      Return lunarMonthNames(_lunarMonth - 1) & "月"
    End Get
  End Property

  ''' <summary>
  ''' 农历日
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarDay() As Integer
    Get
      Return _lunarDay
    End Get
  End Property

  ''' <summary>
  ''' 农历日中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property LunarDayName() As String
    Get
      Dim lunarDayNames As String() = New String() {"初一", "初二", "初三", "初四", "初五", "初六", "初七", "初八", _
        "初九", "初十", "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十", "廿一", "廿二", _
        "廿三", "廿四", "廿五", "廿六", "廿七", "廿八", "廿九", "三十"}
      Return lunarDayNames(_lunarDay - 1)
    End Get
  End Property

  ''' <summary>
  ''' 天干
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HeavenlyStemName() As String
    Get
      Dim heavenlyStemNames As String() = New String() {"甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"}
      Return heavenlyStemNames(_heavenlyStem)
    End Get
  End Property

  ''' <summary>
  ''' 地支
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property EarthlyBranchName() As String
    Get
      Dim earthlyBranchNames As String() = New String() {"子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"}
      Return earthlyBranchNames(_earthlyBranch)
    End Get
  End Property

  ''' <summary>
  ''' 属相(生肖)中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TwelveAnimalName() As String
    Get
      Dim twelveAnimalNames As String() = New String() {"鼠", "牛", "虎", "兔", "龙", "蛇", "马", "羊", "猴", "鸡", "狗", "猪"}
      Return twelveAnimalNames(_earthlyBranch)
    End Get
  End Property

  ''' <summary>
  ''' 节气中文名称
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SolarTermName() As String
    Get
      Dim solarTermNames As String() = New String() {"小寒", "大寒", "立春", "雨水", "惊蛰", "春分", "清明", "谷雨", "立夏", _
        "小满", "芒种", "夏至", "小暑", "大暑", "立秋", "处暑", "白露", "秋分", "寒露", "霜降", "立冬", "小雪", "大雪", "冬至"}
      Return solarTermNames(_solarTerm)
    End Get
  End Property

  ''' <summary>
  ''' 节气日期
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SolarTermDate() As Date
    Get
      Return _solarTermDate
    End Get
  End Property

  ' ''' <summary>
  ' ''' 节日中文名称
  ' ''' </summary>
  ' ''' <value></value>
  ' ''' <returns></returns>
  ' ''' <remarks></remarks>
  'Public ReadOnly Property FestivalNames() As String
  '  Get
  '    Dim fesrivals As New System.Collections.ArrayList
  '    If _isLeapMonth = False Then
  '      Select Case _lunarMonth * 100 + _lunarDay
  '        Case 101
  '          fesrivals.Add("春节")
  '        Case 115
  '          fesrivals.Add("元宵节")
  '        Case 505
  '          fesrivals.Add("端午节")
  '        Case 815
  '          fesrivals.Add("中秋节")
  '        Case 909
  '          fesrivals.Add("重阳节")
  '      End Select
  '    End If
  '    Select Case _solarDate.Month * 100 + _solarDate.Day
  '      Case 101
  '        fesrivals.Add("元旦")
  '      Case 214
  '        fesrivals.Add("情人节")
  '      Case 308
  '        fesrivals.Add("国际妇女节")
  '      Case 315
  '        fesrivals.Add("消费者权益日")
  '      Case 401
  '        fesrivals.Add("愚人节")
  '      Case 501
  '        fesrivals.Add("国际劳动节")
  '      Case 601
  '        fesrivals.Add("国际儿童节")
  '      Case 1001
  '        fesrivals.Add("国庆节")
  '      Case 1225
  '        fesrivals.Add("圣诞节")
  '    End Select
  '    If fesrivals.Count > 0 Then Return String.Join("、", CType(fesrivals.ToArray, String()))
  '    Return String.Empty
  '  End Get
  'End Property

#End Region

#Region "方法"

  ' ****************************************************************************************
  ' 闰月
  ' ****************************************************************************************
  Private Function GetLeapMonth(ByVal solarYear As Integer) As Integer
    Return "0123456789abc".IndexOf(LUNAR_LEAPMONTH.Substring(solarYear - LUNAR_BEGINNINGYEAR, 1))
  End Function

  ' ****************************************************************************************
  ' 节气日期
  ' ****************************************************************************************
  Private Function GetSolarTermDate(ByVal solarYear As Integer, ByVal n As Integer) As Date
    Dim jd As Double = solarYear * (365.2423112 - 0.000000000000064 * (solarYear - 100) * (solarYear - 100) - 0.00000003047 * (solarYear - 100)) + 15.218427 * n + 1721050.71301 '儒略日
    Dim zeta As Double = 0.0003 * solarYear - 0.372781384 - 0.2617913325 * n '角度
    Dim yd As Double = (1.945 * System.Math.Sin(zeta) - 0.01206 * System.Math.Sin(2 * zeta)) * (1.048994 - 0.00002583 * solarYear) ' 年差实均数
    Dim sd As Double = -0.0018 * System.Math.Sin(2.313908653 * solarYear - 0.439822951 - 3.0443 * n) '朔差实均数
    Dim days As Double = jd + yd + sd - 1721425
    Return SpanDays2Date(days)
  End Function

#End Region

#Region "计算农历"

  ''' <summary>
  ''' 计算农历
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CalculateLunar()
    CalculateLunarDay()
    CalculateLunarMonth()
    CalculateLunarYear()
    _leepMonth = GetLeapMonth(_solarDate.Year)
    CalculateSolarTerm()
    _heavenlyStem = ((_lunarYear - 4) Mod 60) Mod 10
    _earthlyBranch = ((_lunarYear - 4) Mod 60) Mod 12
  End Sub

  ''' <summary>
  ''' 计算农历日
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CalculateLunarDay()

    Const rpi As Double = 180 / System.Math.PI
    Const zone As Double = 8.0  '时区:东八区

    Dim t As Double = (_solarDate.Year - 1899.5) / 100.0
    Dim ms As Double = System.Math.Floor((_solarDate.Year - 1900) * 12.3685)
    Dim f0 As Double = Ang(ms, t, 0, 0.75933, 0.0002172, 0.000000155) + 0.53058868 * ms - 0.000837 * t + zone / 24.0 + 0.5
    Dim fc As Double = 0.1734 - 0.000393 * t
    Dim j0 As Double = 693595 + 29 * ms
    Dim aa0 As Double = Ang(ms, t, 0.08084821133, 359.2242 / rpi, 0.0000333 / rpi, 0.00000347 / rpi)
    Dim ab0 As Double = Ang(ms, t, 0.07171366128, 306.0253 / rpi, -0.0107306 / rpi, -0.00001236 / rpi)
    Dim ac0 As Double = Ang(ms, t, 0.08519585128, 21.2964 / rpi, 0.0016528 / rpi, 0.00000239 / rpi)

    Dim aa As Double
    Dim ab As Double
    Dim ac As Double
    Dim f1 As Double
    Dim j As Double
    Dim diff As Integer
    Dim i As Integer

    For i = -1 To 13 'k=整数为朔,k=半整数为望
      aa = aa0 + 0.507984293 * i
      ab = ab0 + 6.73377553 * i
      ac = ac0 + 6.818486628 * i
      f1 = f0 + 1.53058868 * i + fc * System.Math.Sin(aa) - 0.4068 * System.Math.Sin(ab) + 0.0021 * System.Math.Sin(2 * aa) + _
        0.0161 * System.Math.Sin(2 * ab) + 0.0104 * System.Math.Sin(2 * ac) - 0.0074 * System.Math.Sin(aa - ab) - _
        0.0051 * System.Math.Sin(aa + ab)
      j = j0 + 28 * i + f1 '朔或望的等效标准天数及时刻
      diff = Date2SpanDays(_solarDate) - CInt(System.Math.Floor(j)) '当前日距朔日的差值
      If (diff >= 0 AndAlso diff <= 29) Then _lunarDay = diff + 1
    Next

    '下面是对现行农历的校正（1901-2050）,由波波网友指出
    Select Case _solarDate.Date
      Case New Date(1924, 3, 5) To New Date(1924, 4, 3)
        _lunarDay += 1
        If _lunarDay > 30 Then _lunarDay = 1
      Case New Date(2018, 11, 7) To New Date(2018, 12, 6)
        _lunarDay -= 1
        If _lunarDay < 1 Then _lunarDay = 30
      Case New Date(2025, 4, 27) To New Date(2025, 5, 26)
        _lunarDay += 1
        If _lunarDay > 30 Then _lunarDay = 1
    End Select

  End Sub

  ''' <summary>
  ''' 计算农历月
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CalculateLunarMonth()

    Dim leapMonthCount As Integer
    Dim nonLeapMonthCount As Integer
    Dim i As Integer

    For i = 1 To _solarDate.Year
      If (GetLeapMonth(i) > 0) Then
        leapMonthCount += 1
      End If
    Next
    nonLeapMonthCount = Round((_solarDate.Subtract(New Date(1, 2, 12)).Days - _lunarDay) / 29.530588) - leapMonthCount
    '历史上的修改月建
    If (_solarDate.Year <= 240) Then nonLeapMonthCount += 1
    If (_solarDate.Year <= 237) Then nonLeapMonthCount -= 1
    If (_solarDate.Year < 24) Then nonLeapMonthCount += 1
    If (_solarDate.Year < 9) Then nonLeapMonthCount -= 1
    If (_solarDate.Year <= -255) Then nonLeapMonthCount += 1
    If (_solarDate.Year <= -256) Then nonLeapMonthCount += 2
    If (_solarDate.Year <= -722) Then nonLeapMonthCount += 1

    _lunarMonth = nonLeapMonthCount Mod 12 + 1
    If (_lunarMonth = GetLeapMonth(_solarDate.Year - 1) AndAlso _solarDate.Month = 1 AndAlso _solarDate.Day < _lunarDay) Then
      _isLeapMonth = True '如果year-1年末是闰月且该月接到了year年,则year年年初也是闰月
    ElseIf (_lunarMonth = GetLeapMonth(_solarDate.Year)) Then
      If (_solarDate.Month = 1) AndAlso (GetLeapMonth(_solarDate.Year) <> 12) Then
        _lunarMonth += 1 '比如1984年有闰10月，而1984-1-1的lunM=10，但这是从1983年阴历接过来的，所以不是1984年的闰10月
      Else
        _isLeapMonth = True
      End If
    ElseIf (_lunarMonth < GetLeapMonth(_solarDate.Year) OrElse ((_solarDate.Month < _lunarMonth) AndAlso (GetLeapMonth(_solarDate.Year) > 0))) Then
      If _lunarMonth = 12 Then _lunarMonth = 1 Else _lunarMonth += 1 '如果year年是闰月但当月未过闰月则前面多扣除了本年的闰月，这里应当补偿
    End If

    '下面是对现行农历的校正（1901-2050）,由波波网友指出
    Select Case _solarDate.Date
      Case New Date(1924, 3, 5)
        _lunarMonth += 1
      Case New Date(2018, 11, 7)
        _lunarMonth -= 1
      Case New Date(2025, 4, 27)
        _lunarMonth += 1
    End Select

  End Sub

  ''' <summary>
  ''' 计算农历年
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CalculateLunarYear()
    If (_lunarMonth >= 10) AndAlso (_solarDate < New Date(_solarDate.Year, 4, 1)) Then
      _lunarYear = _solarDate.Year - 1
    Else
      _lunarYear = _solarDate.Year
    End If
  End Sub

  ''' <summary>
  ''' 计算节气
  ''' </summary>
  ''' <remarks></remarks>
  Private Sub CalculateSolarTerm()
    Dim date1 As Date
    Dim date2 As Date
    _solarTerm = (_solarDate.Month - 1) * 2
    date1 = GetSolarTermDate(_solarDate.Year, _solarTerm)
    For i As Integer = 1 To 3
      _solarTerm += 1
      If (_solarTerm <= 24) Then
        date2 = GetSolarTermDate(_solarDate.Year, _solarTerm)
      Else
        date2 = GetSolarTermDate(_solarDate.Year + 1, 1)
      End If
      If (_solarDate.Date >= date1.Date) AndAlso (_solarDate.Date < date2.Date) Then
        If (_solarTerm = 1) Then
          _solarTerm = 23
        Else
          _solarTerm -= 2
        End If
        _solarTermDate = date1
        Exit For
      End If
      date1 = date2
    Next
  End Sub

#End Region

#Region "私有函数"

  ''' <summary>
  ''' 返回小数部分
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function Tail(ByVal value As Double) As Double
    Return value - System.Math.Floor(value)
  End Function

  ''' <summary>
  ''' 四舍五入
  ''' </summary>
  ''' <param name="value"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function Round(ByVal value As Double) As Integer
    Return CInt(System.Math.Floor(value + 0.5))
  End Function

  ''' <summary>
  ''' 角度函数
  ''' </summary>
  ''' <param name="x"></param>
  ''' <param name="t"></param>
  ''' <param name="c1"></param>
  ''' <param name="t0"></param>
  ''' <param name="t2"></param>
  ''' <param name="t3"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function Ang(ByVal x As Double, ByVal t As Double, ByVal c1 As Double, ByVal t0 As Double, _
    ByVal t2 As Double, ByVal t3 As Double) As Double
    Return Tail(c1 * x) * 2 * System.Math.PI + t0 - t2 * t * t - t3 * t * t * t
  End Function

  ''' <summary>
  ''' 从1年1月1日起经过的天数
  ''' </summary>
  ''' <param name="solarDate"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function Date2SpanDays(ByVal solarDate As Date) As Integer
    Return solarDate.Subtract(Date.MinValue).Days + 1
  End Function

  ''' <summary>
  ''' 从1年1月1日起经过指定天数的日期
  ''' </summary>
  ''' <param name="days"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SpanDays2Date(ByVal days As Double) As Date
    Return Date.MinValue.AddDays(days - 1)
  End Function

  ''' <summary>
  ''' 从本年1月1日起经过的天数
  ''' </summary>
  ''' <param name="solarDate"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function Date2SpanDaysInYear(ByVal solarDate As Date) As Integer
    Return solarDate.Subtract(New Date(solarDate.Year, 1, 1)).Days + 1
  End Function

  ''' <summary>
  ''' 从本年经过指定天数的日期
  ''' </summary>
  ''' <param name="solarYear"></param>
  ''' <param name="days"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SpanDays2DateInYear(ByVal solarYear As Integer, ByVal days As Double) As Date
    Return (New Date(solarYear, 1, 1)).AddDays(days - 1)
  End Function

#End Region

End Class
