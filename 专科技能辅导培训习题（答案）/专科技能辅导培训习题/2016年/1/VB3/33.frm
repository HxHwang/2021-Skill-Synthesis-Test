VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   8580
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算 "
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "整数n"
      Height          =   735
      Left            =   2520
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "整数m"
      Height          =   735
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim m, n As Integer
Dim sum As Integer


m = Val(Text1.Text)
n = Val(Text2.Text)
sum = 0
If m < n Then
  For i = m To n
     If i Mod 2 <> 0 Then sum = sum + i
  Next i
Else
  For i = m To n Step -1
     If i Mod 2 <> 0 Then sum = sum + i
  Next i
End If

Print sum
     
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub
