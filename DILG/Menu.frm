VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Department of Interior and Local Government"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9160
   ScaleMode       =   0  'User
   ScaleWidth      =   18898.27
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "E&XIT"
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PAYROLL &PROCESS"
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EMPLOYEE'S &RECORD"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   4440
      Picture         =   "Menu.frx":27A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Employee.Show
Me.Hide

End Sub

Private Sub Command2_Click()
Me.Hide
Salary.Show

End Sub

Private Sub Command3_Click()
Main.Show
Unload Me

End Sub

Private Sub Command4_Click()
Me.Enabled = False
Administrator.Show
End Sub

Private Sub Form_Terminate()
Main.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.Show
End Sub
