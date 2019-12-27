VERSION 5.00
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form frmProgress 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   ClientHeight    =   852
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7956
   LinkTopic       =   "Form1"
   ScaleHeight     =   852
   ScaleWidth      =   7956
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin vbalProgBarLib6.vbalProgressBar pb 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   4695
      _ExtentX        =   8276
      _ExtentY        =   656
      Picture         =   "frmProgress.frx":0000
      ForeColor       =   12582912
      BarPicture      =   "frmProgress.frx":001C
      ShowText        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C00000&
      Height          =   600
      Left            =   120
      Top             =   120
      Width           =   6105
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
  Shape1.Move 0, 0, ScaleWidth, ScaleHeight
  pb.Move 120, (ScaleHeight - pb.Height) / 2, ScaleWidth - 240
End Sub
