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
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   615
      Left            =   1800
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "整数n"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "整数m"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim m As Integer
    Dim n As Integer
    Dim s As Long
    Dim i As Integer
    s = 1
    m = Text1.Text
    n = Text2.Text
    For i = m To n
        If i Mod 2 = 0 Then
            s = s * i
        End If
    Next i
    Print s
End Sub
