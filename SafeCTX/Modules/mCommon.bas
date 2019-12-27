Attribute VB_Name = "mCommon"
Option Explicit

Public Sub About()
    Dim Msg As String
    Dim Ver As String
    Dim Hdr As String

    Hdr = "SafeWorX ActiveX Controls"
    Ver = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    Msg = ""
    Msg = Msg & App.CompanyName & vbCrLf & vbCrLf
    Msg = Msg & App.Comments & vbCrLf & vbCrLf
    Msg = Msg & App.FileDescription & vbCrLf & vbCrLf
    Msg = Msg & "©" & App.LegalCopyright & vbCrLf
    
    Msg = Replace(Msg, ".", "." & vbCrLf)
    MsgBox Msg, vbOKOnly, "About SafeWorX Controls"
End Sub


