VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IAutoComplete Demo"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Update Terms"
      Height          =   270
      Left            =   210
      TabIndex        =   16
      Top             =   2100
      Width           =   1290
   End
   Begin VB.Timer Timer1 
      Left            =   3405
      Top             =   3555
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Disable and release"
      Height          =   315
      Left            =   1320
      TabIndex        =   15
      Top             =   3555
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Disable"
      Height          =   315
      Left            =   0
      TabIndex        =   14
      Top             =   3555
      Width           =   1200
   End
   Begin VB.CommandButton Command6 
      Caption         =   "All (including custom)"
      Height          =   315
      Left            =   1965
      TabIndex        =   11
      Top             =   1890
      Width           =   1845
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Enable Custom"
      Height          =   270
      Left            =   210
      TabIndex        =   10
      Top             =   1815
      Width           =   1290
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MRU"
      Height          =   315
      Left            =   1965
      TabIndex        =   9
      Top             =   1500
      Width           =   1230
   End
   Begin VB.CommandButton Command3 
      Caption         =   "History"
      Height          =   315
      Left            =   1965
      TabIndex        =   8
      Top             =   1155
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "File System w/ Options"
      Height          =   420
      Left            =   1965
      TabIndex        =   6
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Basic filesystem"
      Height          =   315
      Left            =   1965
      TabIndex        =   4
      Top             =   285
      Width           =   1245
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   2625
      Width           =   4500
   End
   Begin VB.TextBox Text1 
      Height          =   1470
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   315
      Width           =   1755
   End
   Begin VB.Label Label6 
      Caption         =   "Select only one command from the buttons above. Before switching to another, reset it."
      Height          =   420
      Left            =   15
      TabIndex        =   13
      Top             =   3090
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "General options set in code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   2205
      Width           =   2010
   End
   Begin VB.Label Label4 
      Caption         =   "see code to set options"
      Height          =   405
      Left            =   3255
      TabIndex        =   7
      Top             =   645
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   2265
   End
   Begin VB.Label Label3 
      Caption         =   "Enable with:"
      Height          =   285
      Left            =   2175
      TabIndex        =   5
      Top             =   75
      Width           =   1050
   End
   Begin VB.Label Label2 
      Caption         =   "Test Box:"
      Height          =   255
      Left            =   45
      TabIndex        =   3
      Top             =   2385
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Custom terms, 1 per line"
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   1830
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cACL As cAutoComplete
Private sTerms() As String
Private vTerms As Variant

Private Sub Command1_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
cACL.AC_Filesys Text2.hWnd, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND
Timer1.Interval = 3000
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
cACL.AC_ACList2 Text2.hWnd, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND, ACLO_FILESYSDIRS Or ACLO_DESKTOP
End Sub

Private Sub Command3_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
cACL.AC_History Text2.hWnd, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND
End Sub

Private Sub Command4_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
cACL.AC_MRU Text2.hWnd, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND
End Sub

Private Sub Command5_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
sTerms = Split(Text1.Text, vbCrLf)
cACL.AC_Custom Text2.hWnd, sTerms, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND
End Sub

Private Sub Command6_Click()
If (cACL Is Nothing) Then
    Set cACL = New cAutoComplete
End If
vTerms = Split(Text1.Text, vbCrLf)
cACL.AC_Multi Text2.hWnd, ACO_AUTOSUGGEST Or ACO_AUTOAPPEND, ACLO_FILESYSDIRS, True, True, True, True, vTerms

End Sub

Private Sub Command7_Click()
If (cACL Is Nothing) Then Exit Sub
cACL.AC_Disable
End Sub

Private Sub Command8_Click()
If (cACL Is Nothing) Then Exit Sub

cACL.AC_Disable
Set cACL = Nothing
End Sub

Private Sub Command9_Click()
sTerms = Split(Text1.Text, vbCrLf)
cACL.UpdateCustomTerms sTerms
End Sub

Private Sub Timer1_Timer()
Dim l As Long, s As String
cACL.DropdownStatus l, s
Debug.Print "status=" & l & ",str=" & s
End Sub
