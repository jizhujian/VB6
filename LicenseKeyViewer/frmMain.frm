VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "控件许可证浏览器"
   ClientHeight    =   5148
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   4464
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5148
   ScaleWidth      =   4464
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ListBox lstControlComponentName 
      Height          =   4548
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4212
   End
   Begin VB.TextBox txtLicenseKey 
      Height          =   264
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4212
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  
  Dim controlComponentName() As String
  Dim i As Integer
  
  Set Icon = LoadResIcon("ColorsWatch")

  Dim file As New dotNET2COM.file
  controlComponentName = file.ReadAllLines(App.Path & "\ControlList.txt")
  Set file = Nothing

  For i = LBound(controlComponentName) To UBound(controlComponentName)
    lstControlComponentName.AddItem controlComponentName(i)
  Next

End Sub

Private Sub lstControlComponentName_Click()
  On Error Resume Next
  Err.Clear
  Controls.Add lstControlComponentName.Text, "ce"
  If (Err.Number = 0) Then
    Controls.Remove "ce"
  ElseIf (Err.Number <> 731) Then
    txtLicenseKey.Text = Err.Description
    Exit Sub
  End If
  On Error GoTo HERROR1
  txtLicenseKey.Text = Licenses(lstControlComponentName.Text).LicenseKey
  Exit Sub
HERROR1:
  On Error GoTo HERROR2
  txtLicenseKey.Text = Licenses.Add(lstControlComponentName.Text)
  Exit Sub
HERROR2:
  txtLicenseKey.Text = Err.Number & vbCrLf & Err.Description
  Exit Sub
End Sub
