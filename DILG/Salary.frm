VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Salary 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary"
   ClientHeight    =   6900
   ClientLeft      =   1755
   ClientTop       =   2400
   ClientWidth     =   5745
   Icon            =   "Salary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   5745
   Begin VB.CommandButton Command7 
      Caption         =   "&Close Preview"
      Height          =   495
      Left            =   1680
      TabIndex        =   29
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Print Preview"
      Height          =   495
      Left            =   1680
      TabIndex        =   28
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Print Form"
      Height          =   495
      Left            =   2880
      TabIndex        =   27
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox ho1 
      Height          =   375
      Left            =   9810
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "35"
      Top             =   2475
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox dp1 
      Height          =   375
      Left            =   9750
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "250"
      Top             =   1995
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Total"
      Height          =   225
      Left            =   1515
      TabIndex        =   24
      Top             =   4530
      Width           =   1110
   End
   Begin VB.TextBox Tot 
      Height          =   285
      Left            =   1515
      TabIndex        =   22
      Text            =   "0"
      Top             =   4830
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Search"
      Height          =   375
      Left            =   2070
      TabIndex        =   20
      Top             =   120
      Width           =   1410
   End
   Begin VB.TextBox SSS 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "6%"
      Top             =   5310
      Width           =   1095
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   600
      TabIndex        =   18
      Top             =   6240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   6240
      Width           =   1095
   End
   Begin VB.TextBox T 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0"
      Top             =   4035
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Co&mpute"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   3675
      Width           =   1335
   End
   Begin VB.TextBox OD 
      Height          =   285
      Left            =   1500
      TabIndex        =   10
      Text            =   "0"
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox HO 
      Height          =   285
      Left            =   1515
      TabIndex        =   6
      Text            =   "0"
      Top             =   4125
      Width           =   1095
   End
   Begin VB.TextBox DP 
      Height          =   285
      Left            =   1515
      TabIndex        =   5
      Text            =   "0"
      Top             =   3645
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "id"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   8040
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
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
      RecordSource    =   "Record"
      Caption         =   "Employee"
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Total"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   990
      TabIndex        =   23
      Top             =   4905
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "position"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   2790
      Width           =   195
   End
   Begin VB.Label lname2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "firstname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   195
   End
   Begin VB.Label lname 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      DataField       =   "lastname"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   225
      TabIndex        =   16
      Top             =   1005
      Width           =   195
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Net Pay Total"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3720
      TabIndex        =   14
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Other Deductions"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5790
      Width           =   1245
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "SSS"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   1020
      TabIndex        =   9
      Top             =   5385
      Width           =   315
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Hours Overtime"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   4245
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Days Present"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   3765
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "ID NO."
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "LAST NAME"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "FIRST NAME"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Position"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   195
      TabIndex        =   1
      Top             =   2520
      Width           =   555
   End
End
Attribute VB_Name = "Salary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim searcx
Dim book
Dim found
Private Sub cmdC_Click()
DP = 0
HO = 0
OD = 0
T = 0
Unload PrintF
End Sub
Private Sub Command1_Click()
T = (Val(Tot) - Val(Tot * 0.06)) - Val(OD)
End Sub
Private Sub Command2_Click()
Unload Me
Menu.Show
Unload PrintF
End Sub

Private Sub Command3_Click()
Tot = (Val(DP) * Val(dp1)) + (Val(HO) * Val(ho1))
End Sub

Private Sub Command4_Click()
searcx = InputBox("Enter The ID Number", "Searching")
book = Adodc1.Recordset.Bookmark

Adodc1.Recordset.MoveFirst

Do Until Adodc1.Recordset.EOF Or found
    If Adodc1.Recordset.Fields("id") Like searcx Then
        found = True
    Else
        Adodc1.Recordset.MoveNext
    End If
Loop
    
If found = False Then
    MsgBox "Record not found", vbInformation, "Database Function"
    Adodc1.Recordset.Bookmark = book
End If

End Sub

Private Sub Command5_Click()
PrintF.PrintForm
Unload PrintF
End Sub

Private Sub Command6_Click()
PrintF.Show
Command7.Visible = True
End Sub

Private Sub Command7_Click()
Unload PrintF
Command7.Visible = False
End Sub

Private Sub DP_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then DP = 0
If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub Form_Terminate()
Menu.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Menu.Show

End Sub

Private Sub HO_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then HO = 0
If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub
Private Sub OD_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then OD = 0
If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub


