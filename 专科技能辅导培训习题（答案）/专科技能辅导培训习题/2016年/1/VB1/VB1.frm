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
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "转换大写"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "新字符串"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1.Text = UCase("aBcDeF")
End Sub

Private Sub Command2_Click()
    End
End Sub
