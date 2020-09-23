VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Abbeville Police Employee Timeclock"
   ClientHeight    =   8175
   ClientLeft      =   150
   ClientTop       =   420
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   255
      Left            =   3600
      TabIndex        =   35
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox hidorder 
      Height          =   285
      Left            =   5640
      TabIndex        =   33
      Top             =   6480
      Width           =   615
   End
   Begin VB.CheckBox chksup 
      Caption         =   "sup"
      Height          =   195
      Left            =   5640
      TabIndex        =   32
      Top             =   6720
      Width           =   615
   End
   Begin VB.TextBox txthidtime 
      Height          =   285
      Left            =   5640
      TabIndex        =   31
      Top             =   7680
      Width           =   615
   End
   Begin VB.TextBox txthidmin 
      Height          =   285
      Left            =   5640
      TabIndex        =   30
      Top             =   7440
      Width           =   615
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   360
      TabIndex        =   27
      Top             =   7200
      Width           =   4335
      Begin VB.Label Label6 
         Caption         =   "Total Time:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lbltotal 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   28
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.CheckBox chkshowall 
      Caption         =   "Show All"
      Height          =   255
      Left            =   2520
      TabIndex        =   26
      Top             =   2280
      Width           =   975
   End
   Begin VB.Frame Frame3 
      Caption         =   "Times"
      Height          =   1815
      Left            =   360
      TabIndex        =   21
      Top             =   5400
      Width           =   4335
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10-42:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "10-41:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl1042 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   960
         TabIndex        =   23
         Top             =   1080
         Width           =   3135
      End
      Begin VB.Label lbl1041 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.TextBox txtautonum 
      Height          =   285
      Left            =   5640
      TabIndex        =   20
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox txtrecnum 
      Height          =   285
      Left            =   5640
      TabIndex        =   19
      Top             =   7200
      Width           =   615
   End
   Begin VB.ComboBox cmbidnum 
      Height          =   315
      Left            =   5640
      TabIndex        =   18
      Top             =   6960
      Width           =   615
   End
   Begin VB.CheckBox chkflag 
      Caption         =   "Flag Time to Supervisor"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   360
      TabIndex        =   12
      Top             =   4560
      Width           =   4335
      Begin VB.CommandButton Command4 
         Caption         =   "Exit"
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Cmdlogin 
         Caption         =   "Log In"
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2000
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   4335
      Begin MSComctlLib.ListView ListView1 
         Height          =   1575
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2778
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
         NumItems        =   3
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
      End
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "*"
      Height          =   255
      Left            =   4320
      TabIndex        =   9
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "*"
      Height          =   255
      Left            =   4320
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtend 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtstart 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.ComboBox cmbcategory 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.ComboBox cmbemployee 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label5 
      Height          =   135
      Left            =   0
      TabIndex        =   34
      Top             =   8040
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "End Time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Category:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   1455
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuabout 
         Caption         =   "About"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnumin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuopts 
      Caption         =   "Options"
      Begin VB.Menu mnutime 
         Caption         =   "Time Options"
         Visible         =   0   'False
         Begin VB.Menu mnuhours 
            Caption         =   "Hours"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuminutes 
            Caption         =   "Minutes"
         End
         Begin VB.Menu mnuseconds 
            Caption         =   "Seconds"
         End
      End
      Begin VB.Menu mnuflagunder 
         Caption         =   "Flag Under"
         Begin VB.Menu mnunone 
            Caption         =   "Flag None"
         End
         Begin VB.Menu mnu8 
            Caption         =   "&8 Hours"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnu10 
            Caption         =   "10 Hours"
         End
         Begin VB.Menu mnu12 
            Caption         =   "12 Hours"
         End
      End
      Begin VB.Menu mnusort 
         Caption         =   "Sort Order"
         Begin VB.Menu mnuasc 
            Caption         =   "Ascending"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnudesc 
            Caption         =   "Descending"
         End
      End
   End
   Begin VB.Menu mnusup 
      Caption         =   "Supervisor Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuflagged 
         Caption         =   "View Flagged"
      End
      Begin VB.Menu mnutotalhours 
         Caption         =   "View Total Hours"
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemployee 
         Caption         =   "Employee Management"
      End
      Begin VB.Menu mnucategory 
         Caption         =   "Category Management"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnupath 
         Caption         =   "Path to Database"
      End
   End
   Begin VB.Menu mnuhid 
      Caption         =   "Hidden Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuclick 
         Caption         =   "Click Me"
      End
   End
End
Attribute VB_Name = "Form1"
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




Private Sub chkshowall_Click()
Call Load_List

End Sub

Private Sub chksup_Click()
If chksup.Value = Checked Then
mnusup.Visible = True
Else
mnusup.Visible = False
End If
End Sub

Private Sub cmbemployee_Click()
cmbidnum.ListIndex = cmbemployee.ListIndex


End Sub

Private Sub cmdend_Click()
txtend = Now
End Sub

Private Sub cmdEnter_Click()
On Error Resume Next
If txtstart = "" Then
txtstart = Now
Else
End If




sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Set rsTtest = .OpenRecordset("Select * from timecard")
'If rsTtest.RecordCount > 0 Then
'MsgBox "I'm Sorry, The unit you've selected is already at work"
'Exit Sub
'End If

Set rsttime = .OpenRecordset("timecard")
    With rsttime
    .AddNew
    If cmbemployee = "" Then
    MsgBox "Please Select an Employee."
    Exit Sub
    End If
    If cmbcategory = "" Then
    MsgBox "Please Select an Category."
    Exit Sub
    End If
    
    !employee = cmbemployee
    !category = cmbcategory
    !timeon = txtstart
    !idnum = cmbidnum
    .Update
    End With
    

End With
    Call Load_List
    
    
End Sub

Private Sub Cmdlogin_Click()
If Cmdlogin.Caption = "Log Out" Then
mnusup.Visible = False
Cmdlogin.Caption = "Log In"
Else
frmlogin.Show
frmlogin.Text1 = ""
End If

End Sub

Private Sub cmdstart_Click()
txtstart = Now
End Sub

Private Sub Load_Employees()
'On Error Resume Next

sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If

Set db = OpenDatabase(sfile)
With db
cmbemployee.clear
Set RSTemp = .OpenRecordset("select * from Employee")
    With RSTemp
    .MoveFirst
    Do While Not .EOF
        cmbemployee.AddItem .Fields("name")
        cmbidnum.AddItem .Fields("idnum")
        
        
        .MoveNext
    DoEvents
    Loop
  
End With
End With
End Sub
Private Sub Load_Category()
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
cmbcategory.clear
Set RSTreason = .OpenRecordset("select * from Category")
    With RSTreason
    .MoveFirst
    Do While Not .EOF
        cmbcategory.AddItem .Fields("reason")
        .MoveNext
    DoEvents
    Loop
  
End With
End With
End Sub

Private Sub Command1_Click()
txtstart = ""
txtend = ""
cmbemployee = ""
cmbcategory = ""
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
'If txtend = "" Then
'txtend = Now
'Else
'End If




sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("Select * from timecard where autonum = " & txtautonum)

    With rsttime
    .Edit
    If cmbemployee = "" Then
    MsgBox "Please Select an Employee."
    Exit Sub
    End If
    If cmbcategory = "" Then
    MsgBox "Please Select an Category."
    Exit Sub
    End If
   
    !employee = cmbemployee
    !category = cmbcategory
    !timeon = txtstart
    !idnum = cmbidnum
    !timeoff = txtend
    !totalhours = Format(DateDiff("n", txtstart, txtend) / 60, "##.##")
     If chkflag.Value = Checked Then
    !flag = Checked
    Else
    !flag = Unchecked
    End If
    If mnu8.Checked = True And txthidtime < 8 Then
    !flag = Checked
    ElseIf mnu10.Checked = True And txthidtime < 10 Then
    !flag = Checked
    ElseIf mnu12.Checked = True And txthidtime < 12 Then
    !flag = Checked
    ElseIf mnunone.Checked = True Then
        End If
    
    .Update
    End With
    

End With
    Call Load_List
    Call clear
    cmdUpdate.Enabled = False
    cmbemployee.SetFocus
    
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Call Load_Employees
Call Load_Category
Call Load_List
If mnuasc.Checked = True Then
hidorder = "ASC"
ElseIf mnudesc.Checked = True Then
hidorder = "DESC"
End If
End Sub

Private Sub Load_List()
On Error Resume Next

sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
If chkshowall.Value = 1 Then
Set rsttime = .OpenRecordset("select * from timecard order by autonum " & hidorder)
Else
Set rsttime = .OpenRecordset("select * from timecard where timeoff is null order by autonum " & hidorder)
    End If
    
    With rsttime
    If rsttime.RecordCount = 0 Then ListView1.ListItems.clear: Exit Sub
        
    rsttime.MoveFirst
    
    ListView1.ListItems.clear
    
    Do Until rsttime.EOF
       ListView1.ListItems.Add , , rsttime.Fields("idnum")
        ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 1, , .Fields("employee")
         ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 2, , .Fields("autonum")
      ' ListView1.ListItems.Item(ListView1.ListItems.Count).ListSubItems.Add 3, , .Fields("timeon")
        
          
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

Private Sub Label5_Click()
mnuhid.Visible = True

End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
cmdUpdate.Enabled = True
Dim sfile, answer As String

txtautonum = ListView1.SelectedItem.SubItems(2)
If txtautonum = "" Then
Exit Sub
End If
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Debug.Print "here we go"
Set rsTtest = .OpenRecordset("Select * from timecard where autonum =" & txtautonum)
Debug.Print "we set it"
With rsTtest



  .MoveFirst
        Do While Not .EOF
            If ListView1.SelectedItem.SubItems(2) = .Fields("autonum") Then
            Debug.Print "start put"
                  cmbemployee = !employee
                  cmbcategory = !category
                  txtstart = !timeon
                  cmbidnum = !idnum
                  txtend = !timeoff
                  If !flag = True Then
                  chkflag.Value = 1
                  Else
                  chkflag.Value = 0
                  End If
                  
                  lbl1041 = !timeon
                  If !timeoff <> "" Then
                  lbl1042 = !timeoff
                  lbl1042 = Format(lbl1042, "dddd" & " " & "mmmm dd,yyyy" & " " & "hh:mm:ss")
                  Else
                  lbl1042 = "Still At Work"
                  lbltotal.Caption = "Approxitmately " & DateDiff("h", txtstart, Now) & " hours."
                  lbl1042.ForeColor = vbRed
                  
                  End If
                
                  lbl1041 = Format(lbl1041, "dddd" & " " & "mmmm dd,yyyy" & " " & "hh:mm:ss")
                
                
            End If
            .MoveNext
        Loop
         .MoveFirst
   
    Debug.Print "does it at least get here?"
    
           ' .OpenRecordset ("select * from repairtable where custfName = '" & lstlookup.SelectedItem.Text & "'")
            End With
          End With
          If lbl1042 = "Still At Work" Then
      
          
          
          Exit Sub
          Else
          If txtend = "" Then
          txthidmin = DateDiff("n", txtstart, Now)
          lbltotal = Format(txthidmin / 60, "##.00 Hours So Far")
          Else
          txthidmin = DateDiff("n", txtstart, txtend)
          txthidtime = Val(txthidmin) / 60
          End If
          If mnuhours.Checked = True Then
    lbltotal = Format(txthidtime, "##.00 Hours")
          ElseIf mnuminutes.Checked = True Then
    lbltotal = DateDiff("n", txtstart, txtend) & " minutes"
    ElseIf mnuseconds.Checked = True Then
    lbltotal = DateDiff("s", txtstart, txtend) & " seconds"
          End If
 
End If

End Sub
Public Sub update_form()
    Dim sfile, answer As String


sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
Set rsTtest = .OpenRecordset("Select * from timecard")

With rsTtest
    cmbemployee = !employee
    cmbcategory = !category
    txtstart = !timeon
    cmbidnum = !idnum
    txtend = !timeoff
    End With
    End With
End Sub
Public Sub clear()
cmbemployee.clear
cmbcategory.clear
txtstart = ""
txtend = ""

End Sub

Private Sub mnu10_Click()
mnunone.Checked = False
mnu8.Checked = False
mnu10.Checked = True
mnu12.Checked = False
End Sub

Private Sub mnu12_Click()
mnunone.Checked = False
mnu8.Checked = False
mnu10.Checked = False
mnu12.Checked = True

End Sub

Private Sub mnu8_Click()
mnunone.Checked = False
mnu8.Checked = True
mnu10.Checked = False
mnu12.Checked = False
End Sub

Private Sub mnuabout_Click()
Form5.Show

End Sub

Private Sub mnuasc_Click()
mnuasc.Checked = True
mnudesc.Checked = False
hidorder = "ASC"
Call Load_List
End Sub

Private Sub mnucategory_Click()
Form3.Show

End Sub

Private Sub mnuclick_Click()
Form4.Show

End Sub

Private Sub mnudesc_Click()
mnudesc.Checked = True
mnuasc.Checked = False
hidorder = "DESC"
Call Load_List
End Sub

Private Sub mnuemployee_Click()
Form2.Show

End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuflagged_Click()
frmflagged.Show
End Sub

Private Sub mnuhours_Click()

mnuminutes.Checked = False
mnuseconds.Checked = False
mnuhours.Checked = True
End Sub

Private Sub mnumin_Click()
Me.WindowState = 1

End Sub

Private Sub mnuminutes_Click()
mnuminutes.Checked = True
mnuseconds.Checked = False
mnuhours.Checked = False

End Sub

Private Sub mnunone_Click()
mnunone.Checked = True
mnu8.Checked = False
mnu10.Checked = False
mnu12.Checked = False
End Sub

Private Sub mnupath_Click()
frmpath.Show
End Sub

Private Sub mnuseconds_Click()
mnuminutes.Checked = False
mnuseconds.Checked = True
mnuhours.Checked = False
End Sub

Private Sub mnutotalhours_Click()
frmtotal.Show

End Sub
