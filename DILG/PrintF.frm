VERSION 5.00
Begin VB.Form PrintF 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   7770
   ClientTop       =   3015
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   4395
   ScaleMode       =   0  'User
   ScaleWidth      =   8735.488
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   450
      Top             =   7290
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Salary Payslip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4320
      TabIndex        =   16
      Top             =   600
      Width           =   1950
   End
   Begin VB.Line Line11 
      X1              =   8637.337
      X2              =   0
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line10 
      X1              =   8637.337
      X2              =   8637.337
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line9 
      X1              =   8637.337
      X2              =   0
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   4320
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   8480.294
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   8480.294
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8480.294
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8480.294
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   735
      Width           =   420
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1200
      TabIndex        =   14
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label T 
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2400
      TabIndex        =   13
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Net Pay Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   12
      Top             =   3120
      Width           =   1980
   End
   Begin VB.Label ded 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label OT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label sal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   9
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Total Deductions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   8
      Top             =   2040
      Width           =   1470
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Overtime Charge"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   1425
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   2040
      Width           =   540
   End
   Begin VB.Line Line7 
      X1              =   5496.487
      X2              =   5496.487
      Y1              =   1920
      Y2              =   2865
   End
   Begin VB.Line Line6 
      X1              =   2198.595
      X2              =   2198.595
      Y1              =   1920
      Y2              =   2865
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   8480.294
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label fname 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   195
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2640
      TabIndex        =   4
      Top             =   1680
      Width           =   975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Family Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1050
      WordWrap        =   -1  'True
   End
   Begin VB.Label lname 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   165
      WordWrap        =   -1  'True
   End
   Begin VB.Label Date 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Department of Interior and Local Government"
      BeginProperty Font 
         Name            =   "Adobe Caslon Pro"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "PrintF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
PrintF.PrintForm
Command3.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Date.Caption = Now
lname = Salary.lname
fname = Salary.lname2
T = Salary.T
sal = Val(Salary.DP) * Val(Salary.dp1)
ded = Val(Salary.Tot * 0.06) + Val(Salary.OD)
OT = Val(Salary.HO) * Val(Salary.ho1)
End Sub

