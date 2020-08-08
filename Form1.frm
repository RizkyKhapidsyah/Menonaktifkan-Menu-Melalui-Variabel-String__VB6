VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menonaktifkan Menu Melalui Variabel String"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
   Begin VB.Menu mnuInput 
      Caption         =   "Input"
   End
   Begin VB.Menu mnCari 
      Caption         =   "Cari"
   End
   Begin VB.Menu mnuKeluar 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Buat tipe data sendiri untuk menampung data caption 'dan name-nya menu
Private Type arrMenu
  Caption As String
  Nama As String
End Type
'Deklarasi array dinamis bertipe arrMenu
Dim tabMenu() As arrMenu

'Deklarasi variabel bertipe Control
Dim Contrl As Control

Private Sub Form_Load()
Dim i As Integer, Menu1 As String, Menu2 As String
  'Misalkan Menu1 dan Menu2 ditampung dari database
  Menu1 = "mnuInput"
  Menu2 = "mnuKeluar"
  'Untuk setiap control di Form1
  For Each Contrl In Form1.Controls
      'Periksa, jika Menu maka...
      If (TypeOf Contrl Is Menu) Then
         'Alokasi array untuk menu yang ada
         ReDim Preserve tabMenu(i + 1)
         'Tampung caption dan nama menu
         tabMenu(i).Caption = Contrl.Caption
         tabMenu(i).Nama = Contrl.Name
         'Periksa menu mana saja yang akan di-disabled
         If tabMenu(i).Nama = Menu1 Or _
            tabMenu(i).Nama = Menu2 Then
            'Jika ketemu, disable-kan sekarang
            Contrl.Enabled = False
         End If  'Akhir proses menu disabled
      End If  'Akhir pemeriksaan menu di Form1
  Next Contrl 'Ke control berikutnya
  'Bebaskan memori yang digunakan array
  Erase tabMenu
End Sub


