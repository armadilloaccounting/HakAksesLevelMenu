VERSION 5.00
Object = "{677448D6-C83D-11D2-BEF8-525400DFB47A}#1.1#0"; "cTreeOpt6.ocx"
Begin VB.Form frmHakAkses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hak Akses"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   3735
      TabIndex        =   2
      Top             =   5100
      Width           =   855
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   5100
      Width           =   855
   End
   Begin cTreeOpt6.XTreeOpt treeHakUser 
      Height          =   4860
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8573
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Indentation     =   256.251983642578
   End
End
Attribute VB_Name = "frmHakAkses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public operator As String

Private Function getMenuID(ByVal strKode As String) As String
    getMenuID = Right(strKode, Len(strKode) - 1)
End Function

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    Dim hakAkses   As String
    
    Dim nodX        As Object
    Dim nodY        As Object
    
    Dim x           As Long
    Dim y           As Long
    
    Set nodX = treeHakUser.Nodes(1).Child
    For x = 1 To treeHakUser.Nodes(1).Children
        Set nodY = nodX.Child
        
        For y = 1 To nodX.Children
            If treeHakUser.Value(nodY.Index) = OptionTreeCheckFull Or treeHakUser.Value(nodY.Index) = OptionTreeCheckPartial Then
                hakAkses = hakAkses & getMenuID(nodY.Key) & ","
            End If
            
            Set nodY = nodY.Next
        Next y
        
        Set nodX = nodX.Next
    Next x
    
    If Len(hakAkses) > 0 Then hakAkses = Left(hakAkses, Len(hakAkses) - 1)
    
    If Not Len(hakAkses) > 0 Then
        MsgBox "Minimal 1 hak akses untuk '" & UCase(operator) & "' harus dipilih", vbExclamation, "Peringatan"
    
    Else
        strSql = "UPDATE operator SET hak_akses = '" & hakAkses & "' WHERE operator = '" & rep(operator) & "'"
        conn.Execute strSql
        
        MsgBox "Perubahan hak akses '" & UCase(operator) & "' sudah disimpan.", vbInformation, "Informasi"
        Unload Me
    End If
End Sub

