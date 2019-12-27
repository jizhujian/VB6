VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Interaction"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function InputString(ByRef Value As String, Optional ByVal Prompt, Optional ByVal Caption, _
  Optional ByVal AllowEmpty, Optional MaxLength) As Boolean

  Dim frm As New frmInputString
  With frm
    If (Value > "") Then
      .txtValue.Text = Value
    End If
    If Not IsMissing(Prompt) Then
      .lblPrompt.Caption = Prompt
    End If
    If Not IsMissing(Caption) Then
      .Caption = Caption
    End If
    If Not IsMissing(AllowEmpty) Then
      .mblnAllowEmpty = AllowEmpty
    End If
    If Not IsMissing(MaxLength) Then
      .txtValue.MaxLength = MaxLength
    End If
    .Show vbModal
    If .mblnSuccess Then
      Value = Trim(.txtValue.Text)
      InputString = True
    End If
  End With
  Unload frm
  Set frm = Nothing

End Function

Public Function InputDate(ByRef Value As Date, Optional ByVal Prompt, Optional ByVal Caption, _
  Optional ByVal AllowEmpty, Optional ByVal MinValue, Optional ByVal MaxValue) As Boolean

  Dim frm As New frmInputDate
  With frm
    If (Value > CDate(0)) Then
      .dtpDate.Value = Value
    End If
    If Not IsMissing(Prompt) Then
      .lblPrompt.Caption = Prompt
    End If
    If Not IsMissing(Caption) Then
      .Caption = Caption
    End If
    If Not IsMissing(AllowEmpty) Then
      .dtpDate.CheckBox = AllowEmpty
    Else
      .dtpDate.CheckBox = False
    End If
    If Not IsMissing(MinValue) Then
      .mblnCheckMinValue = True
      .mdtmMinValue = MinValue
    End If
    If Not IsMissing(MaxValue) Then
      .mblnCheckMaxValue = True
      .mdtmMaxValue = MaxValue
    End If
    .Show vbModal
    If .mblnSuccess Then
      If Not IsNull(.dtpDate.Value) Then
        Value = .dtpDate.Value
      Else
        Value = CDate(0)
      End If
      InputDate = True
    End If
  End With
  Unload frm
  Set frm = Nothing

End Function

Public Function InputNumeric(ByRef Value As Double, Optional ByVal Prompt, Optional ByVal Caption, _
  Optional ByVal MinValue, Optional ByVal MaxValue) As Boolean

  Dim frm As New frmInputNumeric
  With frm
    If (Value <> 0) Then
      .txtValue.Text = Value
    End If
    If Not IsMissing(Prompt) Then
      .lblPrompt.Caption = Prompt
    End If
    If Not IsMissing(Caption) Then
      .Caption = Caption
    End If
    If Not IsMissing(MinValue) Then
      .mblnCheckMinValue = True
      .mdblMinValue = MinValue
    End If
    If Not IsMissing(MaxValue) Then
      .mblnCheckMaxValue = True
      .mdblMaxValue = MaxValue
    End If
    .Show vbModal
    If .mblnSuccess Then
      Value = Val(.txtValue.Text)
      InputNumeric = True
    End If
  End With
  Unload frm
  Set frm = Nothing

End Function

Public Function InputMultiLineText(ByRef Value As String, Optional ByVal Prompt, Optional ByVal Caption, _
  Optional ByVal AllowEmpty, Optional MaxLength) As Boolean

  Dim frm As New frmInputMultiLineText
  With frm
    If (Value > "") Then
      .txtValue.Text = Value
    End If
    If Not IsMissing(Prompt) Then
      .lblPrompt.Caption = Prompt
    End If
    If Not IsMissing(Caption) Then
      .Caption = Caption
    End If
    If Not IsMissing(AllowEmpty) Then
      .mblnAllowEmpty = AllowEmpty
    End If
    If Not IsMissing(MaxLength) Then
      .txtValue.MaxLength = MaxLength
    End If
    .Show vbModal
    If .mblnSuccess Then
      Value = Trim(.txtValue.Text)
      InputMultiLineText = True
    End If
  End With
  Unload frm
  Set frm = Nothing

End Function


