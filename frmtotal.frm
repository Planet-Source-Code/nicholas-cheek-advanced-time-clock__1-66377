VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmtotal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Total Time"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8295
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Commands"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   3855
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Query Name Only"
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optdate 
      Caption         =   "Use Date Query"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   120
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Query"
      Enabled         =   0   'False
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CommandButton cmddatequery 
         Caption         =   "Go"
         Height          =   255
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin MSMask.MaskEdBox timeoffbox 
         Height          =   255
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox timeonbox 
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "And"
         Height          =   255
         Left            =   4080
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Date Between:"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   315
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.ComboBox CmbName 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8055
      _ExtentX        =   14208
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
   Begin VB.Frame Frame2 
      Caption         =   "Total Hours"
      Height          =   735
      Left            =   4080
      TabIndex        =   12
      Top             =   4080
      Width           =   4095
      Begin VB.Label ttlhours 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmtotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim db As DAO.Database
Dim RSTuser As DAO.Recordset
Dim rsttime As DAO.Recordset

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long



Private Sub cmddatequery_Click()
On Error Resume Next
ttlhours = "0"
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from timecard where employee = '" & CmbName & "' and timeon between #" & timeonbox & "# and #" & timeoffbox & "#")
    
    
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
             ttlhours = Val(ttlhours) + Val(.Fields("totalhours"))
        
          
             rsttime.MoveNext
    Loop
    End With
    Set rsttime = Nothing
    End With
ListView1.Refresh
Call lvAutosizeMax(ListView1)
End Sub

Private Sub Command1_Click()
On Error Resume Next
ttlhours = "0"
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db

Set rsttime = .OpenRecordset("select * from timecard where employee = '" & CmbName & "'")
    
    
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
             ttlhours = Val(ttlhours) + Val(.Fields("totalhours"))
           
          
             rsttime.MoveNext
            
        Loop
    End With
    
    Set rsttime = Nothing
    End With
ListView1.Refresh

Call lvAutosizeMax(ListView1)
End Sub

Private Sub Command2_Click()
On Error Resume Next


      If ListView1.ListItems.Count > 0 Then
         Printer.Font = "Tahoma"
         Printer.FontBold = True
         Printer.FontUnderline = False
         Printer.FontSize = 30
         Printer.Print
 
          Printer.Print Space(6) & "Hourly Report"
         Printer.Print
         Printer.FontSize = 10
         Printer.Print Space(6) & "Total Hours Worked: " & ttlhours.Caption
         Printer.Print
         
   
         Printer.Print Space(6) & "Report Generated: " & Date
         Printer.Print
         Printer.FontUnderline = False
         Printer.FontBold = False
         Printer.Print vbNewLine

         For i = 1 To ListView1.ListItems.Count
            Printer.Print Space(6) & "Number : " & Str$(i)
            Printer.Print Space(6) & "Employee ID: " & ListView1.ListItems(i).Text
            Printer.Print Space(6) & "Employee Name: " & ListView1.ListItems(i).ListSubItems(1).Text
            Printer.Print Space(6) & "Record Number: " & ListView1.ListItems(i).ListSubItems(2).Text
            Printer.Print Space(6) & "Time In: " & ListView1.ListItems(i).ListSubItems(3).Text
           Printer.Print Space(6) & "Time Out: " & ListView1.ListItems(i).ListSubItems(4).Text
           Printer.Print Space(6) & "Total Hours: " & ListView1.ListItems(i).ListSubItems(5).Text
           
            Printer.Print vbNewLine
         Next i
         Printer.EndDoc
      End If
   
End Sub

Private Sub Command3_Click()
Me.Hide

End Sub

Private Sub Form_Load()
Call Load_users
End Sub


Private Sub Load_users()
sfile = ReadIniFile(App.Path & "\settings.ini", "Config", "Path", "")
If sfile = "NONE" Then
sfile = App.Path & "\timeclock.mdb"
End If
Set db = OpenDatabase(sfile)
With db
CmbName.clear
Set RSTuser = .OpenRecordset("select * from employee")
    With RSTuser
    .MoveFirst
    Do While Not .EOF
        CmbName.AddItem .Fields("name")
        .MoveNext
    DoEvents
    Loop
  
End With
End With
End Sub
Private Sub lvAutosizeMax(lv As ListView)
Dim col2adjust As Long
    For col2adjust = 0 To lv.ColumnHeaders.Count - 1
        Call SendMessage(lv.hwnd, LVM_SETCOLUMNWIDTH, col2adjust, ByVal LVSCW_AUTOSIZE_USEHEADER)
    Next
End Sub

Private Sub optdate_Click()
If optdate.Value = True Then
Command1.Visible = False
Frame1.Enabled = True
Frame1.Visible = True
Else
Command1.Visible = True
Frame1.Enabled = False
Frame1.Visible = False
End If

End Sub

Private Sub Option1_Click()
If optdate.Value = True Then
Command1.Visible = False
Frame1.Enabled = True
Frame1.Visible = True
Else
Command1.Visible = True
Frame1.Enabled = False
Frame1.Visible = False
End If
End Sub
