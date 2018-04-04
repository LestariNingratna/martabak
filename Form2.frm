VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Pesanan Anda"
   ClientHeight    =   4155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   4155
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Text            =   "Isi kue terdiri dari:"
      Top             =   1200
      Width           =   4095
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   4095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Text            =   "Kue"
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   3480
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form2()


End Sub

Private Sub Command1_Click()
Form1.Show
Unload Me
End Sub
