VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmOperator 
   Caption         =   "Operator"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSelesai 
      Caption         =   "Selesai"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdHakAkses 
      Caption         =   "Hak Akses"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdPerbaiki 
      Caption         =   "Perbaiki"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid gridOP 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5953
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   0   'False
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmOperator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub aturGrid()
    Dim kolom   As Byte
    
    With gridOP
        kolom = 0
        .TextMatrix(0, kolom) = "No."
        .ColWidth(kolom) = 500
        .FixedAlignment(kolom) = flexAlignCenterCenter
        .ColAlignment(kolom) = flexAlignCenterCenter
        
        kolom = 1
        .TextMatrix(0, kolom) = "Operator"
        .ColWidth(kolom) = 5050
        .FixedAlignment(kolom) = flexAlignCenterCenter
        .ColAlignment(kolom) = flexAlignLeftCenter
        
        kolom = 2
        .TextMatrix(0, kolom) = "Hak Akses"
        .ColWidth(kolom) = 0
        .ColAlignment(kolom) = flexAlignLeftCenter
    End With
End Sub

Private Sub getDataOpertor()
    Dim objOperator As clsOperator
    
    Set objOperator = New clsOperator
    With objOperator
        If .startGetOperator Then
            Do While .getDataOperator
                gridOP.AddItem gridOP.Rows & vbTab & .operator & vbTab & .hakAkses
            Loop
            gridOP_EnterCell
            
        Else
            cmdPerbaiki.Enabled = False
            cmdHapus.Enabled = False
            cmdHakAkses.Enabled = False
        End If
        .endGetData
    End With
    Set objOperator = Nothing
End Sub

Private Function getSelectedMenu(ByVal menuID As Long, ByVal hakAkses As String) As Long
    Dim arrMenuAkses()  As String
    Dim strMenuID       As String
    
    Dim pos             As Long
    
    If Not Len(hakAkses) > 0 Then
        getSelectedMenu = 0
            
    Else
        If InStr(1, hakAkses, ",") > 0 Then
            arrMenuAkses = Split(hakAkses, ",")
            
            If menuID = arrMenuAkses(LBound(arrMenuAkses)) Then
                strMenuID = menuID & ","
                
            ElseIf menuID = arrMenuAkses(UBound(arrMenuAkses)) Then
                strMenuID = "," & menuID
                
            Else
                strMenuID = "," & menuID & ","
            End If
            
        Else
            strMenuID = CStr(menuID)
        End If
        
        pos = InStr(1, hakAkses, strMenuID)
        getSelectedMenu = IIf(pos > 0, 1, 0)
    End If
End Function

Private Sub showMenu(ByVal operator As String, ByVal tree As XTreeOpt)
    Dim rsMenuInduk     As ADODB.Recordset
    Dim rsMenuAnak      As ADODB.Recordset
    
    Dim selectedMenu    As Long
    Dim keyChild        As String
    Dim daftarHakAkses  As String
    
    'ambil hak akses operator
    strSql = "SELECT hak_akses FROM operator WHERE operator = '" & rep(operator) & "'"
    daftarHakAkses = CStr(dbGetValue(strSql, "1,2,3")) '1,2,3 -> hak akses default
    
    With tree
        .Clear
        
        .AddCheck "mnuAll", , "Daftar Menu Program", , True
        
        'menampilkan menu induk
        strSql = "SELECT id, menu_name, menu_caption " & _
                 "FROM menu_induk " & _
                 "ORDER BY id"
        Set rsMenuInduk = openRecordset(strSql)
        If Not rsMenuInduk.EOF Then
            Do While Not rsMenuInduk.EOF
                .AddCheck rsMenuInduk("menu_name").Value, .Nodes("mnuAll"), rsMenuInduk("menu_caption").Value, , True
                
                'menampilkan menu anak
                strSql = "SELECT id, menu_name, menu_caption " & _
                         "FROM menu_anak " & _
                         "WHERE menu_induk_id = " & rsMenuInduk("id").Value & " " & _
                         "ORDER BY id"
                Set rsMenuAnak = openRecordset(strSql)
                If Not rsMenuAnak.EOF Then
                    Do While Not rsMenuAnak.EOF
                        selectedMenu = getSelectedMenu(rsMenuAnak("id").Value, daftarHakAkses)
                        
                        keyChild = "K" & CStr(rsMenuAnak("id").Value)
                        .AddCheck keyChild, .Nodes(rsMenuInduk("menu_name").Value), rsMenuAnak("menu_caption").Value
                        .Value(keyChild) = selectedMenu
                        
                        rsMenuAnak.MoveNext
                    Loop
                End If
                Call closeRecordset(rsMenuAnak)
                
                rsMenuInduk.MoveNext
            Loop
        End If
        Call closeRecordset(rsMenuInduk)
        
        .ExpandAll
        .Nodes(1).Selected = True
    End With
End Sub

Private Sub cmdHakAkses_Click()
    With frmHakAkses
        .operator = gridOP.TextMatrix(gridOP.Row, 1)
        Call showMenu(.operator, .treeHakUser)
        
        .Caption = "Hak Akses : " & UCase(.operator)
        .Show vbModal
    End With
End Sub

Private Sub cmdHapus_Click()
    Dim objOperator As clsOperator
    
    On Error GoTo errHandle
    
    If MsgBox("Apakah Anda yakin ?", vbExclamation + vbYesNo, "Konfirmasi") = vbYes Then
        Set objOperator = New clsOperator
        objOperator.operator = gridOP.TextMatrix(gridOP.Row, 1)
        If objOperator.delOperator Then gridOP.RemoveItem gridOP.Row
        Set objOperator = Nothing
    End If
    
    Exit Sub
errHandle:
    Select Case Err.Number
        Case 30015: gridOP.Rows = 1
        Case Else: Call msgWarning(Err.Description)
    End Select
End Sub

Private Sub cmdPerbaiki_Click()
    With frmAddEditOperator
        .dataBaru = False
        
        .txtOperator.Text = gridOP.TextMatrix(gridOP.Row, 1)
        .txtOperator.Enabled = False
        .Show vbModal
    End With
End Sub

Private Sub cmdSelesai_Click()
    Unload Me
End Sub

Private Sub cmdTambah_Click()
    With frmAddEditOperator
        .dataBaru = True
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    Call aturGrid
    Call getDataOpertor
End Sub

Private Sub gridOP_EnterCell()
    If gridOP.Rows > 1 Then
        If LCase(gridOP.TextMatrix(gridOP.Row, 1)) = "admin" Then
            cmdHapus.Enabled = False
            cmdHakAkses.Enabled = False

        Else
            cmdHapus.Enabled = True
            cmdHakAkses.Enabled = True
        End If
    End If
End Sub
