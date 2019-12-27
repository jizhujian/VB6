VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRTF 
   Caption         =   "RTF文本转换"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7875
   StartUpPosition =   3  '窗口缺省
   Begin RichTextLib.RichTextBox rtfTextBox 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9128
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRTF.frx":0000
   End
End
Attribute VB_Name = "frmRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

