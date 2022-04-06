VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command2 
      Caption         =   "Çå¿Õ"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "×ó±ß2¸ö×Ö·û"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "ÐÂ×Ö·û´®"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = Left("aBcDeF", 2)
End Sub

Private Sub Command2_Click()
    Text1.Text = ""
End Sub
