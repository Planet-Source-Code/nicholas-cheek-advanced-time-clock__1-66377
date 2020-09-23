VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Management"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Employee Management"
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.TextBox hidid 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2415
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Employee Name"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Employee List:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Emp. Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "ID Number:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rsttime As DAO.Recordset
Dim rstemployee As DAO.Recordset

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Sub Command1_Click()
On Error Resume Next
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Set rstemployee = .OpenRecordset("Select * from employee where idnum = '" & txtid & "'")
If rstemployee.RecordCount > 0 Then
MsgBox "That Employee ID already exists in the Database", vbCritical, "Error"
Exit Sub
Else
Set rstemployee = .OpenRecordset("employee")
With rstemployee
.AddNew
!idnum = txtid
!Name = txtname
.Update
End With

End If
End With
Call Load_List
txtid = ""
txtname = ""

End Sub

Private Sub Command2_Click()
On Error Resume Next
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from employee where id = " & hidid)
If rsttime.RecordCount = 0 Then
MsgBox "Cannot update a record that does not exist"
Exit Sub

Else

With rsttime
.Edit
!idnum = txtid
!Name = txtname
.Update
End With

End If
End With
Call Load_List

End Sub

Private Sub Command3_Click()
On Error Resume Next
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from employee where id = " & hidid)
If rsttime.RecordCount = 0 Then
MsgBox "Cannot update a record that does not exist"
Exit Sub

Else

With rsttime
.Delete

End With

End If
Call Load_List
End With
txtid = ""
txtname = ""
End Sub

Private Sub Form_Load()
Call Load_List

End Sub



Private Sub Load_List()
On Error Resume Next

sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from employee order by idnum")
    
    With rsttime
    If rsttime.RecordCount = 0 Then ListView1.ListItems.clear: Exit Sub
        
    rsttime.MoveFirst
    
    ListView1.ListItems.clear
    
    Do Until rsttime.EOF
       ListView1.ListItems.Add , , rsttime.Fields("idnum")
        ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 1, , .Fields("name")
        
          
             rsttime.MoveNext
    Loop
    End With
    Set rsttime = Nothing
    End With
ListView1.Refresh
Call lvAutosizeMax(ListView1)
End Sub

Private Sub lvAutosizeMax(lv As ListView)
Dim col2adjust As Long
    For col2adjust = 0 To lv.ColumnHeaders.Count - 1
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next

Dim sfile, answer As String
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Debug.Print "here we go"
Set rstemployee = .OpenRecordset("Select * from employee")
Debug.Print "we set it"
With rstemployee



  .MoveFirst
        Do While Not .EOF
            If ListView1.SelectedItem = .Fields("idnum") Then
            Debug.Print "start put"
                  txtid = !idnum
                  txtname = !Name
                  hidid = !id
                  
                
            End If
            .MoveNext
        Loop
         .MoveFirst
   
    Debug.Print "does it at least get here?"
    
           ' .OpenRecordset ("select * from repairtable where custfName = '" & lstlookup.SelectedItem.Text & "'")
            End With
          End With
        
 


End Sub
