VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Menu 
   Caption         =   " Menu Utama Program Akuntansi PT XXX  ***   "
   ClientHeight    =   2865
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Menu.frx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1560
      Top             =   1440
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":B8EB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":B930A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":B975C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":B9BAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BA000
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BA452
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BA8A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BACF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BB148
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BB59A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BB9EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":BBE3E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   2064
      ButtonWidth     =   2223
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F1 Simpan Kas"
            Key             =   "F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F2 Perkiraan"
            Key             =   "F2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F3 Pemesan"
            Key             =   "F3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F4 Order"
            Key             =   "F4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F5 Pengiriman"
            Key             =   "F5"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F6 Lap Order"
            Key             =   "F6"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F7 Lap Arus Kas"
            Key             =   "F7"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F8 Lap Biaya"
            Key             =   "F8"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "F9 Buku Besar"
            Key             =   "F9"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Keluar ESC"
            Key             =   "ESC"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin Crystal.CrystalReport CR 
      Left            =   840
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Menu mnfile 
      Caption         =   "File Master"
      Begin VB.Menu mnsimpankas 
         Caption         =   "Simpan Kas"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnperkiraan 
         Caption         =   "Perkiraan"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mntransaksi 
      Caption         =   "Transaksi"
      Begin VB.Menu mnPemesan 
         Caption         =   "Pemesan"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnorder 
         Caption         =   "Order"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnpengiriman 
         Caption         =   "Pengiriman"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnlaporan 
      Caption         =   "Laporan"
      Begin VB.Menu mnlaporder 
         Caption         =   "Order"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnlaparuskas 
         Caption         =   "Arus Kas"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnlapbiaya 
         Caption         =   "Biaya - Biaya"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnbukubesar 
         Caption         =   "Buku Besar"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnkeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Bergerak As Integer
Dim Teks As String

Private Sub Form_Load()
Teks = Me.Caption
End Sub

Private Sub mnPemesan_Click()
Pemesan.Show vbModal
End Sub

Private Sub Timer1_Timer()
Me.Caption = Bergerak
Teks = Right(Teks, Len(Teks) - 1) & Left(Teks, 1)
Me.Caption = Teks
End Sub

Private Sub Form_KeyPress(Keyascii As Integer)
If Keyascii = 27 Then
    Pesan = MsgBox("Yakin akan keluar dari program ini..?", vbYesNo)
    If Pesan = vbYes Then
        Call ImplodeForm(Me, 1000)
        End
    End If
End If
End Sub

Private Sub mnbukubesar_Click()
    CR.ReportFileName = App.Path & "\Lap Buku Besar.rpt"
    CR.WindowState = crptMaximized
    CR.RetrieveDataFiles
    CR.Action = 1
End Sub

Private Sub mnkeluar_Click()
Pesan = MsgBox("Yakin akan mengakhiri program ini", vbYesNo)
If Pesan = vbYes Then End
End Sub

Private Sub mnlaparuskas_Click()
LapArusKas.Show vbModal
End Sub

Private Sub mnlapbiaya_Click()
LapBiaya.Show vbModal
End Sub

Private Sub mnlaporder_Click()
LapOrder.Show vbModal
End Sub

Private Sub mnorder_Click()
Order.Show vbModal
End Sub

Private Sub mnsimpankas_Click()
Kas.Show vbModal
End Sub

Private Sub mnpengiriman_Click()
Pengiriman.Show vbModal
End Sub

Private Sub mnperkiraan_Click()
Perkiraan.Show vbModal
End Sub

Sub ceksaldo()
Call BukaDB
RSKas.Open "select * from kas where saldo>0", Conn
RSKas.Requery
If RSKas.EOF Then
    Pesan = MsgBox("kas masih kosong, simpan uang kas dulu", vbYesNo)
    If Pesan = vbYes Then
        Kas.Show
    Else
        End
    End If
End If

End Sub

Private Sub mnsql_Click()
UjiSQL.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "F1"
        Kas.Show vbModal
    Case "F2"
        Perkiraan.Show vbModal
    Case "F3"
        Pemesan.Show vbModal
    Case "F4"
        Order.Show vbModal
    Case "F5"
        Pengiriman.Show vbModal
    Case "F6"
        LapOrder.Show vbModal
    Case "F7"
        LapArusKas.Show vbModal
    Case "F8"
        LapBiaya.Show vbModal
    Case "F9"
        CR.ReportFileName = App.Path & "\Lap Buku Besar.rpt"
        CR.WindowState = crptMaximized
        CR.RetrieveDataFiles
        CR.Action = 1
    Case "ESC"
        Pesan = MsgBox("Yakin akan keluar dari program ini..?", vbYesNo)
        If Pesan = vbYes Then
            Call ImplodeForm(Me, 5000)
            End
        End If
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call ImplodeForm(Me, 1000)
End Sub
