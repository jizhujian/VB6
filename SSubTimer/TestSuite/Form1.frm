VERSION 5.00
Begin VB.Form frmMultiClientSubClassTest 
   Caption         =   "Multi-Client SubClass Tester"
   ClientHeight    =   2925
   ClientLeft      =   1980
   ClientTop       =   2535
   ClientWidth     =   6345
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   6345
   Begin VB.CheckBox chkClass2 
      Caption         =   "Class &2 Enabled"
      Height          =   375
      Left            =   660
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkClass1 
      Caption         =   "Class &1 Enabled"
      Height          =   375
      Left            =   660
      TabIndex        =   0
      Top             =   300
      Value           =   1  'Checked
      Width           =   2415
   End
End
Attribute VB_Name = "frmMultiClientSubClassTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_c1 As cSubclass1
Private m_c2 As cSubclass2

Private Sub chkClass1_Click()
   If (chkClass1.Value = vbChecked) Then
      m_c1.Attach Me.hWnd
   Else
      m_c1.Detach
   End If
End Sub

Private Sub chkClass2_Click()
   If (chkClass2.Value = vbChecked) Then
      m_c2.Attach Me.hWnd
   Else
      m_c2.Detach
   End If
End Sub

Private Sub Form_Load()
   Set m_c1 = New cSubclass1
   m_c1.Attach Me.hWnd
   Set m_c2 = New cSubclass2
   m_c2.Attach Me.hWnd
End Sub
