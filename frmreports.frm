VERSION 5.00
Begin VB.Form frmcategory 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Category"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Add Category"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   255
         Left            =   2040
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtcat 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmcategory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rstcat As DAO.Recordset

Private Sub Command1_Click()
On Error Resume Next
If txtcat = "" Then
MsgBox "Please Enter a Category Title"
Exit Sub
Else


sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Set rstcat = .OpenRecordset("category")
With rstcat
.AddNew
!reason = txtcat
.Update
End With
End With
End If
End Sub

