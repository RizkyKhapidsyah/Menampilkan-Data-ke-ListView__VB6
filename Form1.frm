VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   Caption         =   "Menampilkan Data ke ListView"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim i As Integer
  x = 0
  With ListView1
    .View = lvwReport  'Buat tampilan report...
    'Tambahkan 3 kolom...
    .ColumnHeaders.Add , , "Kolom ke-1"
    .ColumnHeaders.Add , , "Kolom ke-2"
    .ColumnHeaders.Add , , "Kolom ke-3"
    'Tambahkan data sebanyak 20...
    For i = 1 To 20
      .ListItems.Add 1, Key:="", Text:="Data 1 ke-" & i
      .ListItems(1).ListSubItems.Add , , _
       "Data 2 ke-" & i
      .ListItems(1).ListSubItems.Add , , _
       "Data 3 ke-" & i
    Next i
  End With
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader _
As MSComctlLib.ColumnHeader)
'Jika header kolom diklik, data akan disortir secara 'Ascending atau Descending
Select Case ColumnHeader
       Case "Kolom ke-1"
            If ListView1.SortOrder = lvwDescending Then
               ListView1.SortOrder = lvwAscending
            Else
               ListView1.SortOrder = lvwDescending
            End If
            ListView1.Sorted = True
       Case "Kolom ke-2"
            If ListView1.SortOrder = lvwDescending Then
               ListView1.SortOrder = lvwAscending
            Else
               ListView1.SortOrder = lvwDescending
            End If
            ListView1.Sorted = True
       Case "Kolom ke-3"
            If ListView1.SortOrder = lvwDescending Then
               ListView1.SortOrder = lvwAscending
            Else
               ListView1.SortOrder = lvwDescending
            End If
            ListView1.Sorted = True
End Select
End Sub



