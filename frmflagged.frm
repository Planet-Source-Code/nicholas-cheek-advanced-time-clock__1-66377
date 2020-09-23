VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmflagged 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flagged Calls"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8460
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4895
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Employee"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Record Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Time In"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time Out"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmflagged"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Declare variables
Public db As DAO.Database
Public RSTemp As DAO.Recordset
Public RSTreason As DAO.Recordset
Public rsttime As DAO.Recordset
Public rsTtest As DAO.Recordset
Public sfile As String


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Sub Load_List()
On Error Resume Next

sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from timecard where flag = true")
    
    
    With rsttime
    If rsttime.RecordCount = 0 Then ListView1.ListItems.clear: Exit Sub
        
    rsttime.MoveFirst
    
    ListView1.ListItems.clear
    
    Do Until rsttime.EOF
       ListView1.ListItems.Add , , rsttime.Fields("idnum")
        ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 1, , .Fields("employee")
         ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 2, , .Fields("autonum")
          ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 3, , .Fields("timeon")
           ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 4, , .Fields("timeoff")
            ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 5, , .Fields("totalhours")
        
        
          
             rsttime.MoveNext
    Loop
    End With
    Set rsttime = Nothing
    End With
ListView1.Refresh
Call lvAutosizeMax(ListView1)
End Sub

Private Sub Form_Load()
Call Load_List

End Sub
Private Sub lvAutosizeMax(lv As ListView)
Dim col2adjust As Long
    For col2adjust = 0 To lv.ColumnHeaders.Count - 1
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next

sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from timecard where autonum = " & ListView1.SelectedItem.SubItems(2))
    With rsttime
    .Edit
    
    !flag = "False"
    
    .Update
    
    With rsttime
    End With
    End With
    End With
    Call Load_List
    
End Sub
