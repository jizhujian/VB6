Imports System.Net.Mail
Imports System.Text.Encoding

''' <summary>
''' 发送邮件
''' </summary>
''' <remarks></remarks>
<Microsoft.VisualBasic.ComClass(MailSender.ClassId, MailSender.InterfaceId, MailSender.EventsId)> _
Public Class MailSender

    ''' <summary>
    ''' COM注册必须
    ''' </summary>
    ''' <remarks></remarks>
    Public Const ClassId As String = "9b14e36d-30bc-48aa-a5bb-12ac792fa740"
    Public Const InterfaceId As String = "040e7309-b4a1-4e06-aa57-81a213f6ee42"
    Public Const EventsId As String = "4f9d11c7-1354-4c31-911f-7581d11133c2"

    Private _address As String
    Private _displayName As String
    Private _host As String
    Private _port As Integer = 25
    Private _userName As String
    Private _password As String

    Public Sub New()
        MyBase.New()
    End Sub

    ''' <summary>
    ''' Email地址
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Address As String
        Get
            Return _address
        End Get
        Set(value As String)
            _address = value
        End Set
    End Property

    ''' <summary>
    ''' 显示名称
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property DisplayName As String
        Get
            Return _displayName
        End Get
        Set(value As String)
            _displayName = value
        End Set
    End Property

    ''' <summary>
    ''' SMTP发送服务器
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Host As String
        Get
            Return _host
        End Get
        Set(value As String)
            _host = value
        End Set
    End Property

    ''' <summary>
    ''' SMTP发送服务器端口号，默认25
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property Port As Integer
        Get
            Return _port
        End Get
        Set(value As Integer)
            _port = value
        End Set
    End Property

    ''' <summary>
    ''' 帐号
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property UserName As String
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property

    ''' <summary>
    ''' 密码
    ''' </summary>
    ''' <value></value>
    ''' <remarks></remarks>
    Public WriteOnly Property Password As String
        Set(value As String)
            _password = value
        End Set
    End Property

    ''' <summary>
    ''' 发送邮件
    ''' </summary>
    ''' <param name="toList">收件人，多个收件人用分号间隔</param>
    ''' <param name="ccList">抄送，多个抄送用分号间隔</param>
    ''' <param name="bccList">密送，多个密送用分号间隔</param>
    ''' <param name="title">主题</param>
    ''' <param name="content">正文</param>
    ''' <remarks></remarks>
    Public Sub SendMail(ByVal toList As String, ByVal ccList As String, ByVal bccList As String, ByVal title As String, ByVal content As String)

        Dim message As New MailMessage
        '发件人
        message.From = New MailAddress(_address, _displayName)
        '收件人
        If (toList IsNot Nothing) AndAlso (toList.Trim > "") Then
            Dim toArray As String() = toList.Split(";")
            For Each [to] In toArray
                If [to].Trim > "" Then message.To.Add([to])
            Next
        End If
        '抄送
        If (ccList IsNot Nothing) AndAlso (ccList.Trim > "") Then
            Dim ccArray As String() = ccList.Split(";")
            For Each cc In ccArray
                If cc.Trim > "" Then message.CC.Add(cc)
            Next
        End If
        '密送
        If (bccList IsNot Nothing) AndAlso (bccList.Trim > "") Then
            Dim bccArray As String() = bccList.Split(";")
            For Each bcc In bccArray
                If bcc.Trim > "" Then message.Bcc.Add(bcc)
            Next
        End If
        '主题
        message.Subject = title
        message.SubjectEncoding = UTF8
        '正文
        message.Body = content
        message.BodyEncoding = UTF8
        message.IsBodyHtml = True

        Dim o As New SmtpClient
        o.Host = Host
        o.Port = Port
        o.Credentials = New System.Net.NetworkCredential(_userName, _password)
        o.Send(message)

    End Sub

End Class
