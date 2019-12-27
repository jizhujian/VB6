VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBDataReader"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const mstrModuleName As String = "VBCommon.DBDataReader"
Private mobjLog As New Log

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBOpenConnection(ByVal strConnectionString As String) As ADODB.Connection
  Dim cnn As New ADODB.Connection
  cnn.Open strConnectionString
  cnn.CommandTimeout = 0
  Set DBOpenConnection = cnn
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get DBReadCommittedSQL() As String
  DBReadCommittedSQL = "SET NOCOUNT ON;" & vbCrLf & _
    "SET TRANSACTION ISOLATION LEVEL READ COMMITTED;"
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Property Get DBReadUnCommittedSQL() As String
  DBReadUnCommittedSQL = "SET NOCOUNT ON;" & vbCrLf & _
    "SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED;"
End Property

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBAutoExecuteReader(ByVal strConnectString As String, ByVal sql As String, Optional ByVal blnSingleRow As Boolean) As Collection
  Dim cnn As ADODB.Connection
  Set cnn = DBOpenConnection(strConnectString)
  On Error GoTo HERROR
  Set DBAutoExecuteReader = DBExecuteReader(cnn, sql, blnSingleRow)
  On Error GoTo 0
  cnn.Close
  Set cnn = Nothing
  Exit Function
HERROR:
  cnn.Close
  Set cnn = Nothing
  Err.Raise Err.Number, mobjLog.GenErrSource(mstrModuleName, "DBAutoExecuteReader")
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBExecuteReader(ByVal cnn As ADODB.Connection, ByVal sql As String, Optional ByVal blnSingleRow As Boolean) As Collection
  Dim rs As ADODB.Recordset
  Set rs = cnn.Execute(sql, , adCmdText)
  Set DBExecuteReader = DBPopulateRecordSet(rs, blnSingleRow)
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBPopulateRecordSet(ByVal rs As ADODB.Recordset, Optional ByVal blnSingleRow As Boolean, Optional ByVal blnCloseRecordset As Boolean = True) As Collection
  Dim colRecordset As Collection
  If (rs.State = 1) Then
    If (Not rs.EOF) Then
      If blnSingleRow Then
        Set colRecordset = DBPopulateRecord(rs)
      Else
        Set colRecordset = New Collection
        Do
          If rs.EOF Then Exit Do
          colRecordset.Add DBPopulateRecord(rs)
          rs.MoveNext
        Loop
      End If
    End If
    If blnCloseRecordset Then rs.Close
  End If
  Set rs = Nothing
  Set DBPopulateRecordSet = colRecordset
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Private Function DBPopulateRecord(ByVal rs As ADODB.Recordset) As Collection
  Dim colRecordset As Collection
  Dim fld As ADODB.Field
  Set colRecordset = New Collection
  For Each fld In rs.Fields
    colRecordset.Add rs(fld.Name).Value, fld.Name
  Next
  Set DBPopulateRecord = colRecordset
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBAutoExecuteScalar(ByVal strConnectionString As String, ByVal sql As String) As Variant
  Dim cnn As ADODB.Connection
  Set cnn = DBOpenConnection(strConnectionString)
  On Error GoTo HERROR
  DBAutoExecuteScalar = DBExecuteScalar(cnn, sql)
  On Error GoTo 0
  cnn.Close
  Set cnn = Nothing
  Exit Function
HERROR:
  cnn.Close
  Set cnn = Nothing
  Err.Raise Err.Number, mobjLog.GenErrSource(mstrModuleName, "DBAutoExecuteScalar")
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBExecuteScalar(ByVal cnn As ADODB.Connection, ByVal sql As String) As Variant
  Dim rs As ADODB.Recordset
  Set rs = cnn.Execute(sql, , adCmdText)
  DBExecuteScalar = DBPopulateScalar(rs)
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBPopulateScalar(ByVal rs As ADODB.Recordset, Optional ByVal blnCloseRecordset As Boolean = True) As Variant
  If (rs.State = 1) Then
    If (Not rs.EOF) Then
      DBPopulateScalar = rs(0).Value
    End If
    If blnCloseRecordset Then rs.Close
  End If
  Set rs = Nothing
End Function

' ***********************************************************************************************
'
' ***********************************************************************************************
Public Function DBAutoExecuteLines(ByVal strConnectString As String, ByVal sql As String) As String
  Dim cnn As ADODB.Connection
  Set cnn = DBOpenConnection(strConnectString)
  On Error GoTo HERROR
  DBAutoExecuteLines = DBExecuteLines(cnn, sql)
  On Error GoTo 0
  cnn.Close
  Set cnn = Nothing
  Exit Function
HERROR:
  cnn.Close
  Set cnn = Nothing
  Err.Raise Err.Number, mobjLog.GenErrSource(mstrModuleName, "DBAutoExecuteLines")
End Function

Public Function DBExecuteLines(ByVal cnn As ADODB.Connection, ByVal sql As String) As String
  Dim rs As ADODB.Recordset
  Set rs = cnn.Execute(sql, , adCmdText)
  DBExecuteLines = DBPopulateLines(rs)
End Function

Public Function DBPopulateLines(ByVal rs As ADODB.Recordset, Optional ByVal blnCloseRecordset As Boolean = True) As String
  Dim objStringBuilder As New dotNET2COM.StringBuilder
  If (rs.State = 1) Then
    Do
      If rs.EOF Then Exit Do
      objStringBuilder.AppendLine rs(0).Value
      rs.MoveNext
    Loop
    If blnCloseRecordset Then rs.Close
  End If
  DBPopulateLines = objStringBuilder.ToString
End Function

Public Function DBCloneRecordset(ByVal rs As ADODB.Recordset) As ADODB.Recordset
  Dim fld As ADODB.Field
  Dim rsNew As ADODB.Recordset
  Set rsNew = New ADODB.Recordset
  rsNew.CursorLocation = adUseClient
  For Each fld In rs.Fields
    rsNew.Fields.Append fld.Name, fld.Type, fld.DefinedSize, fld.Attributes
    Select Case fld.Type
      Case adNumeric, adDecimal
        With rsNew.Fields(fld.Name)
          .Precision = fld.Precision
          .NumericScale = fld.NumericScale
        End With
    End Select
  Next
  rsNew.Open
  If Not rs.EOF Then
    Do
      If rs.EOF Then Exit Do
      rsNew.AddNew
      For Each fld In rs.Fields
        rsNew(fld.Name).Value = fld.Value
      Next
      rsNew.Update
      rs.MoveNext
    Loop
    rsNew.MoveFirst
  End If
  Set DBCloneRecordset = rsNew
End Function
