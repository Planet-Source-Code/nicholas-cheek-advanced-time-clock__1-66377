VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmpath 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Database Path"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change"
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   330
      Width           =   4095
   End
End
Attribute VB_Name = "frmpath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1 = OpenCommonDialog("Select your database file.", "Timeclock|*.mdb", "*.mde")

End Sub

Private Sub Command2_Click()
Unload Me


End Sub

Private Sub Form_Load()
Text1 = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Text1.Text = "" Then
WriteIniFile App.Path & "\settings.ini", "Config", "Path", "NONE"
Else
WriteIniFile App.Path & "\settings.ini", "Config", "Path", Text1.Text
End If
End Sub
