VERSION 5.00
Begin VB.Form frmlogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Admin Login"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.CommandButton Command3 
         Caption         =   "Exit"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Log In"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1440
         Width           =   2535
      End
      Begin VB.ComboBox cmbuser 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Administration Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public db As DAO.Database
Public RSTuser As DAO.Recordset


Private Sub Load_users()
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
cmbuser.clear
Set RSTuser = .OpenRecordset("select * from users")
    With RSTuser
    .MoveFirst
    Do While Not .EOF
        cmbuser.AddItem .Fields("username")
        .MoveNext
    DoEvents
    Loop
  
End With
End With
End Sub

Private Sub Command1_Click()
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Set RSTuser = .OpenRecordset("select * from users where username = '" & cmbuser.Text & "'")
    With RSTuser
If Text1.Text = !Password Then
Form1.mnusup.Visible = True
Me.Hide
Form1.Cmdlogin.Caption = "Log Out"
Else
Exit Sub
End If
End With
End With


End Sub

Private Sub Form_Load()
Call Load_users

End Sub
