VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Employee 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employees Record"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   Icon            =   "Employee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox gender 
      DataField       =   "gender"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   9720
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
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
      RecordSource    =   "Record"
      Caption         =   "Employee's Record"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Employee Add / Update Record"
      ForeColor       =   &H80000007&
      Height          =   3855
      Left            =   210
      TabIndex        =   6
      Top             =   4035
      Width           =   6375
      Begin VB.TextBox Text7 
         DataField       =   "contact"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   3360
         Width           =   3135
      End
      Begin VB.TextBox Text6 
         DataField       =   "address"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   645
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   20
         Top             =   2520
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H8000000E&
         Caption         =   "Female"
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   2220
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H8000000E&
         Caption         =   "Male"
         Enabled         =   0   'False
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   285
         TabIndex        =   18
         Top             =   2220
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         DataField       =   "position"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   17
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text4 
         DataField       =   "dob"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         DataField       =   "firstname"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox Text2 
         DataField       =   "lastname"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         DataField       =   "id"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   345
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "C&ancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdS 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "CONTACT"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   3360
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ADDRESS"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2520
         Width           =   780
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Position"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "Date of Birth"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1440
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "FIRST NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "LAST NAME"
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000E&
         Caption         =   "ID NO."
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "E&XIT"
      Height          =   375
      Left            =   6720
      TabIndex        =   5
      Top             =   3480
      Width           =   2895
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "&DELETE"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdU 
      Caption         =   "&UPDATE"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&ADD"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3480
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Employee.frx":27A2
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5741
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   6960
      Picture         =   "Employee.frx":27B7
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2520
   End
End
Attribute VB_Name = "Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mb


Private Sub cmdD_Click()

Adodc1.Recordset.delete
If Adodc1.Recordset.RecordCount = 0 Then
cmdD.Enabled = False
Else
cmdD.Enabled = True
End If
End Sub

Private Sub cmdS_Click()
Option1.Value = False
Option2.Value = False
cmdX.Enabled = True
cmdU.Enabled = True
cmdD.Enabled = True
Command1.Enabled = True
cmdS.Enabled = True
Command6.Enabled = True
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
Text6.Enabled = False
Text7.Enabled = False
Option1.Enabled = False
Option2.Enabled = False
If Len(Text1) > 0 And Len(Text2) > 0 And Len(Text3) > 0 And Len(Text4) > 0 And Len(Text5) > 0 And Len(Text6) > 0 And Len(Text7) > 0 Then
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete
Else
    mb = MsgBox("Invalid Data", vbCritical, "KSK")
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
Option1.Value = False
Option2.Value = False
cmdX.Enabled = False
cmdU.Enabled = False
cmdD.Enabled = False
Command1.Enabled = False
End If
End Sub

Private Sub cmdU_Click()
Option1.Value = False
Option2.Value = False
cmdX.Enabled = False
cmdU.Enabled = False
cmdD.Enabled = False
cmdS.Enabled = True
Command6.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
addupdate
End Sub

Private Sub cmdX_Click()
Employee.Hide
Menu.Show
End Sub
Private Sub Command1_Click()
Option1.Value = False
Option2.Value = False
cmdX.Enabled = False
cmdU.Enabled = False
cmdD.Enabled = False
Adodc1.Recordset.AddNew
Text1 = "DILG" & Round(Rnd() * 99) & Chr(Round(Rnd() * 25) + 65)
Text1.SetFocus
cmdS.Enabled = True
Command6.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Option1.Enabled = True
Option2.Enabled = True
addupdate
End Sub


Private Sub Command6_Click()
Unload Me
Employee.Show
End Sub
Private Function delete()
DataGrid1.Refresh

If Adodc1.Recordset.RecordCount = 0 Then
    Adodc1.Enabled = False
    DataGrid1.Enabled = False
    
    cmdD.Enabled = False
    cmdU.Enabled = False
Else
    Adodc1.Enabled = True
    DataGrid1.Enabled = True
    
    cmdD.Enabled = True
    cmdU.Enabled = True
End If

End Function

Private Sub Form_Load()
Option1.Value = False
Option2.Value = False
If Adodc1.Recordset.RecordCount = 0 Then
cmdD.Enabled = False
Else
cmdD.Enabled = True
End If
End Sub

Private Sub Form_Terminate()
Main.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
Adodc1.Recordset.CancelBatch
Main.Show
End Sub
Private Function savecancel()
DataGrid1.Refresh

Command1.Enabled = True
cmdS.Enabled = False
Command6.Enabled = False

End Function
Private Function addupdate()
Command1.Enabled = False
cmdU.Enabled = False
cmdS.Enabled = True
Command6.Enabled = True
cmdD.Enabled = False

Adodc1.Enabled = False
DataGrid1.Enabled = False
End Function

Private Sub Option1_Click()
If Option1.Value = True Then
gender.Text = "Male"
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
gender.Text = "Female"
End If
End Sub

