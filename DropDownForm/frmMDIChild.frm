VERSION 5.00
Begin VB.Form frmMDIChild 
   Caption         =   "Document"
   ClientHeight    =   2880
   ClientLeft      =   4740
   ClientTop       =   3195
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMDIChild.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   2880
   ScaleWidth      =   5730
   Begin VB.TextBox txtMain 
      Height          =   2775
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMDIChild.frx":014A
      Top             =   60
      Width           =   5595
   End
End
Attribute VB_Name = "frmMDIChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   On Error Resume Next
   txtMain.Move 2 * Screen.TwipsPerPixelX, 2 * Screen.TwipsPerPixelY, Me.ScaleWidth - 4 * Screen.TwipsPerPixelX, Me.ScaleHeight - 4 * Screen.TwipsPerPixelY
End Sub
