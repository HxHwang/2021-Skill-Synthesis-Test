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
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command2 
      Caption         =   "���"
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ת��Сд"
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "���ַ���"
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
    Label2.Caption = LCase("aBcDeF")
End Sub

Private Sub Command2_Click()
    Label2.Caption = ""
End Sub
