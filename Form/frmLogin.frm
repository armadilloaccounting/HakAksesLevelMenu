VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   930
      Width           =   855
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   930
      Width           =   855
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2055
   End
   Begin VB.TextBox txtOperator 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Operator"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub disableMenu(ByVal fMain As Form, ByVal hakAkses As String)
    Dim ctl     As Control
    Dim rsMenu  As ADODB.Recordset
    
    strSql = "SELECT menu_name " & _
             "FROM menu_anak " & _
             "WHERE id NOT IN (" & hakAkses & ") " & _
             "ORDER BY id"
    Set rsMenu = openRecordset(strSql)
    If Not rsMenu.EOF Then
        Do While Not rsMenu.EOF
            For Each ctl In fMain.Controls
                If TypeName(ctl) = "Menu" Then
                    If ctl.Name = rsMenu("menu_name").Value Then
                        ctl.Enabled = False
                        Exit For
                    End If
                End If
            Next
            
            rsMenu.MoveNext
        Loop
    End If
    Call closeRecordset(rsMenu)
End Sub

Private Sub cmdLogin_Click()
    Dim objOperator     As clsOperator
    Dim statusOperator  As STATUS_OPERATOR
    Dim hakAkses        As String
    
    If isEmptyText(txtOperator, "Operator") Then Exit Sub
    If isEmptyText(txtPassword, "Password") Then Exit Sub
    
    Set objOperator = New clsOperator
    With objOperator
        .operator = txtOperator.Text
        .password = txtPassword.Text
        
        statusOperator = .isValidUser
        hakAkses = .hakAkses
    End With
    Set objOperator = Nothing
    
    Select Case statusOperator
        Case OP_TDK_DITEMUKAN
            Call msgWarning("Operator belum terdaftar !!!")
            txtOperator.SetFocus
            
        Case OP_PASS_SALAH
            Call msgWarning("Password salah")
            txtPassword.SetFocus
            
        Case OP_PASS_OK
            Call disableMenu(frmMain, hakAkses) 'panggil prosedur disableMenu disini
            frmMain.Show
            
            Unload Me
    End Select
End Sub

Private Sub cmdBatal_Click()
    Unload Me
End Sub

