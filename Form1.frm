VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Pesanan Kue"
   ClientHeight    =   6885
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   6855
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   3600
      TabIndex        =   13
      Top             =   3480
      Width           =   2655
      Begin VB.OptionButton Option7 
         Caption         =   "Dibawa Pulang"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Dimakan di tempat"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ketebalan"
      Height          =   1695
      Left            =   360
      TabIndex        =   10
      Top             =   3360
      Width           =   2775
      Begin VB.OptionButton Option5 
         Caption         =   "Tipis"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tebal"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ukuran"
      Height          =   2655
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2895
      Begin VB.OptionButton Option3 
         Caption         =   "Kecil"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sedang"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Besar"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton CmdSelesai 
      Caption         =   "Selesai"
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Top             =   5520
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buat Kue"
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Keju"
      Height          =   735
      Left            =   3720
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Kacang"
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Coklat"
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Isi Kue"
      Height          =   195
      Left            =   3840
      TabIndex        =   0
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSelesai_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Form1.Option1.Value = True And Form1.Option4.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option1.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If

If Form1.Option1.Value = True And Form1.Option4.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option1.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If

If Form1.Option1.Value = True And Form1.Option5.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option1.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If


If Form1.Option1.Value = True And Form1.Option5.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option1.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If

If Form1.Option2.Value = True And Form1.Option4.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option2.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If

If Form1.Option2.Value = True And Form1.Option4.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option2.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If

If Form1.Option2.Value = True And Form1.Option5.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option2.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If


If Form1.Option2.Value = True And Form1.Option5.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option2.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If

If Form1.Option3.Value = True And Form1.Option4.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option3.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If

If Form1.Option3.Value = True And Form1.Option4.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option3.Caption + " " + Form1.Option4.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If

If Form1.Option3.Value = True And Form1.Option5.Value = True And Form1.Option6.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option3.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option6.Caption
End If


If Form1.Option3.Value = True And Form1.Option5.Value = True And Form1.Option7.Value = True Then
Form2.Text1.Text = "Ukuran " + Form1.Option3.Caption + " " + Form1.Option5.Caption
Form2.Text6.Text = Form1.Option7.Caption
End If


If Form1.Check1.Value = 1 And Form1.Check2.Value = 0 And Form1.Check3.Value = 0 Then
Form2.Text2.Text = Form1.Check1.Caption
End If

If Form1.Check1.Value = 1 And Form1.Check2.Value = 1 And Form1.Check3.Value = 0 Then
Form2.Text2.Text = Form1.Check1.Caption
Form2.Text4.Text = Form1.Check2.Caption
End If

If Form1.Check1.Value = 1 And Form1.Check2.Value = 1 And Form1.Check3.Value = 1 Then
Form2.Text2.Text = Form1.Check1.Caption
Form2.Text4.Text = Form1.Check2.Caption
Form2.Text5.Text = Form1.Check3.Caption
End If

If Form1.Check1.Value = 0 And Form1.Check2.Value = 1 And Form1.Check3.Value = 0 Then
Form2.Text2.Text = Form1.Check2.Caption
End If

If Form1.Check1.Value = 0 And Form1.Check2.Value = 1 And Form1.Check3.Value = 1 Then
Form2.Text2.Text = Form1.Check2.Caption
Form2.Text4.Text = Form1.Check3.Caption
End If

If Form1.Check1.Value = 0 And Form1.Check2.Value = 0 And Form1.Check3.Value = 1 Then
Form2.Text2.Text = Form1.Check3.Caption
End If

Form2.Show
End Sub



