VERSION 5.00
Begin VB.Form frmAddEditOperator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operator"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   3135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOperator 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   2055
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
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   930
      Width           =   855
   End
   Begin VB.CommandButton cmdBatal 
      Caption         =   "Batal"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   930
      Width           =   855
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   525
      Width           =   690
   End
End
Attribute VB_Name = "frmAddEditOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dataBaru As Boolean
    
Private Sub cmdBatal_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    Dim objOperator As clsOperator
    Dim ret         As Boolean
    Dim hakAksesDef As String
    
    If isEmptyText(txtOperator, "Operator") Then Exit Sub
    
    If dataBaru Then
        If isEmptyText(txtPassword, "Password") Then Exit Sub
        
        Set objOperator = New clsOperator
        With objOperator
            .operator = txtOperator.Text
            .password = txtPassword.Text
            
            ret = .addOperator
            hakAksesDef = .hakAkses
        End With
        Set objOperator = Nothing
        
        If ret Then 'nambah operator berhasil
            With frmOperator.gridOP
                .AddItem .Rows & vbTab & txtOperator.Text & vbTab & hakAksesDef
            End With
            
            txtOperator.Text = ""
            txtPassword.Text = ""
            txtOperator.SetFocus
            
        Else
            Call msgWarning("Operator sudah terdaftar")
            txtOperator.SetFocus
        End If
        
    Else
        Set objOperator = New clsOperator
        With objOperator
            .operator = txtOperator.Text
            .password = txtPassword.Text
            
            ret = .editOperator
        End With
        Set objOperator = Nothing
        
        If ret Then 'edit password operator berhasil
            Unload Me
        End If
    End If
End Sub
