VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Demo Pengaturan Hak Akses Level Menu"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuMaster 
      Caption         =   "Master"
      Begin VB.Menu mnuBarang 
         Caption         =   "Barang"
      End
      Begin VB.Menu mnuSpr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Customer"
      End
      Begin VB.Menu mnuSupplier 
         Caption         =   "Supplier"
      End
   End
   Begin VB.Menu mnuTransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnuPembelian 
         Caption         =   "Pembelian"
      End
      Begin VB.Menu mnuReturPembelian 
         Caption         =   "Retur Pembelian"
      End
      Begin VB.Menu mnuSpr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPenjualan 
         Caption         =   "Retur Penjualan"
      End
   End
   Begin VB.Menu mnuBiaya 
      Caption         =   "Biaya"
      Begin VB.Menu mnuBiayaOperasional 
         Caption         =   "Biaya"
      End
      Begin VB.Menu mnuGajiKaryawan 
         Caption         =   "Gaji Karyawan"
      End
   End
   Begin VB.Menu mnuLaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnuLapPembelian 
         Caption         =   "Laporan Pembelian"
      End
      Begin VB.Menu mnuLapJthTempo 
         Caption         =   "Laporan Jatuh Tempo Pembelian"
      End
      Begin VB.Menu mnuSpr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLapPenjualan 
         Caption         =   "Laporan Penjualan"
      End
   End
   Begin VB.Menu mnuPengaturan 
      Caption         =   "Pengaturan"
      Begin VB.Menu mnuOperator 
         Caption         =   "Operator"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuOperator_Click()
    frmOperator.Show vbModal
End Sub
