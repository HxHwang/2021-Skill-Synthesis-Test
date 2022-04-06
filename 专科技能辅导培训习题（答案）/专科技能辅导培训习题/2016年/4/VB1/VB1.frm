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
      Caption         =   "ÍË³ö"
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÓÒ±ß2¸ö×Ö·û"
      Height          =   615
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "ÐÂ×Ö·û´®"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = Right("aBcDeF", 2)
End Sub

Private Sub Command2_Click()
    End
End Sub
