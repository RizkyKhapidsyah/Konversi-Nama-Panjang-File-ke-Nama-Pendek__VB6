VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengkonversi Nama Panjang File ke Nama Pendek"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function sGetShortFileName(ByVal FileName As String) As String
Dim lRC As Long
Dim sShortPath As String
  sShortPath = String$(PATH_LEN + 1, 0)
  lRC = GetShortPathName(FileName, sShortPath, PATH_LEN)
  sGetShortFileName = Left$(sShortPath, lRC)
End Function

Private Sub Form_Load()
  'Ganti 'c:\program files' dengan path/file yang Anda
  'ingin mengetahui nama pendeknya. Nama path harus
  'sudah ada di PC Anda...
  MsgBox sGetShortFileName("c:\Program Files")
End Sub


