VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRs         As ADODB.Recordset

Private mOperator   As String
Private mPassword   As String
Private mHakAkses   As String

Public Property Let operator(ByVal vData As String)
    mOperator = vData
End Property
Public Property Get operator() As String
    operator = mOperator
End Property

Public Property Let password(ByVal vData As String)
    mPassword = vData
End Property
Public Property Get password() As String
    password = mPassword
End Property

Public Property Let hakAkses(ByVal vData As String)
    mHakAkses = vData
End Property
Public Property Get hakAkses() As String
    hakAkses = mHakAkses
End Property

Public Function isValidUser() As STATUS_OPERATOR
    Dim ret As Long
    
'    On Error GoTo errHandle
    
    strSql = "SELECT COUNT(*) FROM operator WHERE operator = '" & rep(operator) & "'"
    ret = CLng(dbGetValue(strSql, 0))
    If ret > 0 Then
        strSql = "SELECT hak_akses FROM operator WHERE operator = '" & rep(operator) & "' AND password = '" & rep(password) & "'"
        mHakAkses = CStr(dbGetValue(strSql, ""))
        If Len(hakAkses) > 0 Then
            isValidUser = OP_PASS_OK
            
        Else
            isValidUser = OP_PASS_SALAH
        End If
        
    Else
        isValidUser = OP_TDK_DITEMUKAN
    End If
    
    Exit Function
errHandle:
    isValidUser = OP_TDK_DITEMUKAN
End Function

Public Function addOperator() As Boolean
''    On Error GoTo errHandle
    
    mHakAkses = "1,2,3" 'set hak akses default, misal 1,2,3
    
    strSql = "INSERT INTO operator (operator, [password], hak_akses) VALUES ('" & _
             rep(operator) & "', '" & rep(password) & "', '" & hakAkses & "')"
    conn.Execute strSql
    
    addOperator = True
    
    Exit Function
errHandle:
    addOperator = False
End Function

Public Function editOperator() As Boolean
''    On Error GoTo errHandle
                    
    strSql = "UPDATE operator SET [password] = '" & rep(password) & "' " & _
             "WHERE operator = '" & rep(operator) & "'"
    conn.Execute strSql
    
    editOperator = True
    
    Exit Function
errHandle:
    editOperator = False
End Function

Public Function delOperator() As Boolean
''    On Error GoTo errHandle
    
    strSql = "DELETE FROM operator " & _
             "WHERE operator = '" & rep(operator) & "'"
    conn.Execute strSql
    
    delOperator = True
    Exit Function
errHandle:
    delOperator = False
End Function

Public Function startGetOperator() As Boolean
'    On Error GoTo errHandle

    strSql = "SELECT operator, password, hak_akses " & _
             "FROM operator " & _
             "ORDER BY operator"
    Set mRs = openRecordset(strSql)
    If Not mRs.EOF Then startGetOperator = True
    
    Exit Function
errHandle:
    startGetOperator = False
End Function

Public Function getDataOperator() As Boolean
'    On Error Resume Next
        
    If Not mRs.EOF Then
        mOperator = IIf(IsNull(mRs("operator").Value), "", mRs("operator").Value)
        mPassword = IIf(IsNull(mRs("password").Value), "", mRs("password").Value)
        mHakAkses = IIf(IsNull(mRs("hak_akses").Value), "", mRs("hak_akses").Value)
        
        getDataOperator = True
        mRs.MoveNext
    End If
End Function

Public Sub endGetData()
    Call closeRecordset(mRs)
End Sub

