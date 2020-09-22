VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Main 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DILG"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10740
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   10740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpass2 
      DataField       =   "pass"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DILG\Employee.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\DILG\Employee.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Log"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtuser2 
      DataField       =   "user"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox txtpass 
      Alignment       =   2  'Center
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2760
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtuser 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      MaxLength       =   12
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "&Log In"
      Default         =   -1  'True
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   135
      Left            =   5280
      TabIndex        =   6
      Top             =   4800
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   6855
      Left            =   2040
      Picture         =   "Main.frx":27A2
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   6840
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mb
Private Sub cmdEnter_Click()
Adodc1.Refresh
Adodc1.Recordset.Find "user =" & "'" & txtuser & "'"

If txtpass = txtpass2 And Len(txtpass) > 0 Then
    Menu.Show
    Unload Me
    
Else
    mb = MsgBox("Invalid username or password.", vbCritical, "KSK")
    txtuser.SetFocus
    txtuser = ""
    txtpass = ""
    SendKeys "{Home}+{End}"
End If
End Sub
Private Sub Command2_Click()
End
End Sub


Private Sub Label1_DblClick()
MsgBox "Powered by : ALROB WORKS INTERACTIVE"
MsgBox "#047 Dau St.,Monte Maria Vill.,Catalunan Grande,Davao City"
MsgBox "Tel.No.(082)299-8945 E-Mail: alrobworks@yahoo.com  www.alrobworks.wetpaint.com"
MsgBox "We make your imagination, To Reality"
End Sub
