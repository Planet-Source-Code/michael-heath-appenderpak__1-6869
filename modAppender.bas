Attribute VB_Name = "modAppender"
' Module Name:    modAppender
' Purpose    :    Append files and extract files to and from a single file
' Coder(s)   :    Michael Heath
' Acknowlegements:  Tim Butler & Robert Carter

'The following code uses the 'GET' and 'PUT' commands to their
'full capability.  These are used to append files into a single file
'& then extract them with no problem.
'NOTE: Its been said that there is a bug with the 'PUT' command,
'from M$ themselves.    That's why I used the 'lngOffset' variable
' to offset the starting position of each extraction.
'Using this, all the 'JUNK' sort of 'disappears'!
'The result is a successful extraction from an appended file
'There may be a size limit on files.  I haven't checked anything above 1.2 meg
'Code may still have some minor bugs.

Public CurrentFileName As String     ' Holds filename
Public pos As Double                 ' Counts position in file
Public Char As Byte                  ' Binary Access String
Public Char2 As String
Public vCount(1 To 65) As String 'Holds size of each file
Public FileString(1 To 65) As String 'Holds name of each file
Public Position(1 To 65) As String 'Holds the starting position(in bytes) of each file
Public NewPosition As Long
Public CurrentCount As Double 'Last position in appended file
Public vSize As Long 'Size of file
Public Type Extracts ' 32-bit access type
    lByte As Byte
    lmByte As Byte
    hmByte As Byte
    hByte As Byte
End Type
Public ExtractBytes As Extracts
Public Hmany As Long ' Number of files appended
Public lngOffset As Long 'Magic Offset used to extract from the appended file
Public Sub OpenFile(vForm As Form)
' Handle errors
    On Error GoTo OpenProblem
    vForm.CommonDialog1.InitDir = App.Path & "\Save"
    vForm.CommonDialog1.Filter = "ALL Files(*.*|*.*"
    vForm.CommonDialog1.FilterIndex = 1
    ' Display an Open dialog box.
    vForm.CommonDialog1.Action = 1
    
    ' vForm.Caption = strAppName & " - " & vForm.CommonDialog1.FileName
    ' vForm.Caption = strAppName & " - " & vForm.CommonDialog1.FileName
    CurrentFileName = vForm.CommonDialog1.FileName
    Exit Sub
OpenProblem:
    ' Cancel button clicked
    Exit Sub

End Sub

Public Sub SaveFile(vForm As Form)
On Error GoTo SaveERR
    Dim FileNum As Integer
    ' Set Initial Directory to open and FileTypes
        vForm.CommonDialog1.InitDir = App.Path & "\Save"
        vForm.CommonDialog1.Filter = "ALL Files | *.*"
        
        CurrentFileName = FileString(frmMain.lstFiles.ListIndex)
        If CurrentFileName = "" Then
            vForm.CommonDialog1.FileName = FileString(frmMain.lstFiles.ListIndex)
        Else
            vForm.CommonDialog1.FileName = CurrentFileName
        End If
        
            vForm.CommonDialog1.ShowSave
            CurrentFileName = vForm.CommonDialog1.FileName
Exit Sub
SaveERR:
        ' I don't know what the error was, but I want to let you know and then
        ' Exit the sub.
End Sub

Public Sub GetBinSize()
' Operation Will Not be stopped once it is this far unless there is an error
Hmany = Hmany + 1
On Error GoTo BinERR
vSize = FileLen(CurrentFileName)
    
' Put the size of file in an index
vCount(Hmany) = vSize
If Hmany = 1 Then
CurrentCount = 1
Else
CurrentCount = CurrentCount + vCount(Hmany - 1)
End If
Exit Sub
BinERR:
    MsgBox "The Following Error Occurred:  " & Err.Description
    Hmany = Hmany - 1
End Sub
Public Sub SaveFileInPak()
Dim TempPos As Long
Dim FileIndex As Integer
Dim pos As Long
' Name the file to be placed in the pak file, operation cannot be canceled
frmMain.ProgressBar1.Min = 1
frmMain.ProgressBar1.Max = vCount(Hmany) + 1

NameFile:
    FileString(Hmany) = InputBox("Please enter the name and file extension for this file", "NameFile")
        If FileString(Hmany) = "" Then
        MsgBox "You can not cancel this operation, you must name the file", vbOKOnly + vbCritical, "Error"
        GoTo NameFile
        End If
' The position of the file in the pak file will be the previous files length + 1
        If Hmany = 1 Then
            NewPosition = 0
            Position(1) = CurrentCount + lngOffset
        ElseIf Hmany >= 2 Then
            FileIndex = Hmany - 1
            NewPosition = vCount(FileIndex)
            lngOffset = lngOffset + 2
            Position(Hmany) = CurrentCount + lngOffset
        End If
' Save the pak list
        SavePakList
' code to actually add the new file to the pak file
Dim strTemp As String
strTemp = PutFileInString(CurrentFileName)
Open App.Path & "\tmhb.apd" For Append As #2
        Print #2, strTemp
        Close #2
        MsgBox "Done"
        frmMain.lstFiles.AddItem FileString(Hmany) & ":                 Size - " & vCount(Hmany) & " Bytes"

End Sub
Public Sub SavePakList()
Dim i As Integer
        Open App.Path & "\paklist.lst" For Output As #3
            Write #3, Hmany, CurrentCount, lngOffset
            For i = 1 To Hmany
            Write #3, FileString(i), vCount(i), Position(i)
            Next i
        Close #3

End Sub
Public Sub ReadPakList()
Dim i As Integer
lngOffset = -4 'Init to the magic first value
frmMain.lstFiles.AddItem "**** Current Files ****"
    Open App.Path & "\paklist.lst" For Input As #1
        Input #1, Hmany, CurrentCount, lngOffset
        MsgBox "No of files= " & Hmany
        For i = 1 To Hmany
        Input #1, FileString(i), vCount(i), Position(i)
        frmMain.lstFiles.AddItem FileString(i) & "       Size - " & vCount(i)
        Next i
    Close #1
    
End Sub
Public Sub SimpleExtract()
'On Error GoTo ExtractERR
Dim i As Long
Dim pos As Long
Dim pos2 As Long
    frmMain.ProgressBar1.Min = 1
    frmMain.ProgressBar1.Max = vCount(frmMain.lstFiles.ListIndex)
    frmMain.Refresh

pos = Position(frmMain.lstFiles.ListIndex)
pos2 = -3
Open CurrentFileName For Binary As #1 ' Was said it couldn't be done?
Open App.Path & "\tmhb.apd" For Binary As #2
    For i = 1 To vCount(frmMain.lstFiles.ListIndex) Step 4
    pos = pos + 4
    pos2 = pos2 + 4
    DoEvents
    frmMain.ProgressBar1.Value = i
    Get #2, pos, ExtractBytes 'Retrieve Data from append file
    Put #1, pos2, ExtractBytes 'Save data to new file

    Next i
    Close #1
    Close #2
    MsgBox "Done"
    Exit Sub
ExtractERR:
    MsgBox Err.Description
    
End Sub
Public Function PutFileInString(sFileName As String) As String
    'sFileName must include Path and file na
    '     me
    'eg "c:\Windows\notepad.exe"
    Dim iFree As Integer, sizeOfFile As Long
    Dim sFileString As String, sTemp As String
    iFree = FreeFile
    Open sFileName For Binary Access Read As iFree
    sizeOfFile = LOF(iFree)
    sFileString = Space$(sizeOfFile)
    Get iFree, , sFileString
    Close #iFree
    PutFileInString = sFileString
    
    ' Thanks to Robert Carter for this function
End Function
'DeleteAll was used for debug purposes
Public Function DeleteAll()
On Error Resume Next
    Dim i As Integer
    Kill App.Path & "\tmhb.apd"
    Kill App.Path & "\paklist.lst"
    frmMain.lstFiles.Clear
    
    For i = 1 To 65
        Position(i) = ""
        vCount(i) = ""
        FileString(i) = ""
    Next
    CurrentCount = 1
    Hmany = 0
    
    ReadPakList
End Function
