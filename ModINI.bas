Attribute VB_Name = "modINI"

'This is how you write to the ini file
'WriteIniFile App.Path & "\GDB.ini", "Misc", "Last-Time-In", Time
'This is how you read the ini file, it's putting it to a textbox, you can sub the text
'box for a label or something like that
'Text8.Text = ReadIniFile(App.Path & "\GDB.ini", "Database", "Path-to-database", "")
   
'================================================================

Option Explicit

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
      "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
      String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias _
      "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
      lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

''***********************************************************************
''***********************************************************************
''***********************************************************************
'Purpose:   To write to an .ini file for a give Selection.
'Arguments: IniFileName - the name of the ini file (including .ini).
'                          if the path is not included, windows will put
'                          the file into the windows directory.
'           sSection -     name of the section heading (don't include [])
'           sItem -        The item heading (the item before the =)
'           sText -        text to write to the ini file.
'Returns:   True/False
'Uses:      WritePrivateProfileString
'***********************************************************************
Function WriteIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sText As String) As Boolean
   Dim i As Integer
   On Error GoTo sWriteIniFileError

   i = WritePrivateProfileString(sSection, sItem, sText, sIniFileName)
   WriteIniFile = True

   Exit Function
sWriteIniFileError:
   WriteIniFile = False
End Function

''***********************************************************************
''***********************************************************************
''***********************************************************************
'Purpose:   Will read an .ini file for a given selection.
'Arguments: sIniFileName - the name of the ini file (including .ini).
'                          if the path is not included, windows will put
'                          the file into the windows directory.
'           sSection -     name of the section heading (don't include [])
'           sItem -        The item heading (the item before the =)
'           sDefault -     If the item, section, or the file doesn't
'                          exist, then this will be returned.
'Returns:   the value of sItem, or sDefault if not found
'Uses:      GetPrivateProfileString()
'***********************************************************************'
Function ReadIniFile(ByVal sIniFileName As String, ByVal sSection As String, ByVal sItem As String, ByVal sDefault As String) As String
   Dim iRetAmount As Integer   'the amount of characters returned
   Dim sTemp As String

   sTemp = String$(200, 0)   'fill with nulls
   iRetAmount = GetPrivateProfileString(sSection, sItem, sDefault, sTemp, 200, sIniFileName)
   sTemp = Left$(sTemp, iRetAmount)
   ReadIniFile = sTemp
End Function
'***********************************************************************
'***********************************************************************

