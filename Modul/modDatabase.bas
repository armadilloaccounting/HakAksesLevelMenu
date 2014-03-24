Attribute VB_Name = "modDatabase"
Option Explicit

Public conn    As ADODB.Connection
Public strSql  As String

Public Function openDb() As Boolean
    Dim strConn As String
    
    On Error GoTo errHandle
            
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sampleDB.mdb"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = strConn
    conn.Open
    
    openDb = True
    
    Exit Function

errHandle:
    openDb = False
End Function

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset
    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    Set openRecordset = obj
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then vRs.Close
    End If
    
    Set vRs = Nothing
End Sub

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue  As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set rsDbGetValue = openRecordset(query)
    If Not rsDbGetValue.EOF Then
        If Not IsNull(rsDbGetValue(0).Value) Then
            dbGetValue = rsDbGetValue(0).Value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If
        
    Call closeRecordset(rsDbGetValue)
    
    Exit Function
errHandle:
    dbGetValue = defValue
End Function
