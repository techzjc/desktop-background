VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "������"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3705
   LinkTopic       =   "Form2"
   ScaleHeight     =   5280
   ScaleWidth      =   3705
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton Command5 
      Caption         =   "����(&L)"
      Height          =   855
      Left            =   600
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "����(&E)"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ָ�����(&R)"
      Height          =   855
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���Ƴ���(&C)"
      Height          =   855
      Left            =   600
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���г���(&R)"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = Shell("D:\1.bat")
End Sub

Private Sub Command2_Click()
x = Shell("D:\2.bat")
MsgBox "���Ƴɹ���"
End Sub

Private Sub Command3_Click()
x = Shell("D:\3.bat")
End Sub

Private Sub Command4_Click()
End
Unload Me
End Sub

Private Sub Command5_Click()
Load Form1
Form1.Show
Unload Me
End Sub
