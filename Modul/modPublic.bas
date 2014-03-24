Attribute VB_Name = "modPublic"
Option Explicit

Public Function isEmptyText(ByVal obj As Object, ByVal title As String) As Boolean
    If Not Len(obj.Text) > 0 Then
        isEmptyText = True
        
        MsgBox "Maaf, informasi '" & title & "' harus diisi !", vbExclamation, "Peringatan"
        obj.SetFocus
    End If
End Function

Public Function rep(ByVal kata As String) As String
    rep = Replace(kata, "'", "''")
End Function

Public Sub msgWarning(ByVal prompt As String)
    MsgBox prompt, vbExclamation, "Warning"
End Sub

