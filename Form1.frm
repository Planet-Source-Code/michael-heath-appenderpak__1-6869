VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simple Pak Creator - No Compression"
   ClientHeight    =   5070
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   5565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   750
      TabIndex        =   2
      Top             =   -300
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4980
      Top             =   3930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstFiles 
      Height          =   4350
      Left            =   30
      TabIndex        =   1
      Top             =   90
      Width           =   5505
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   345
      Left            =   90
      TabIndex        =   0
      Top             =   4560
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "&Add File"
      End
      Begin VB.Menu mnuFileExtract 
         Caption         =   "&Extract File"
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuHelpCredits 
         Caption         =   "&Credits"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
  Frame1.Width = Screen.Width + 100
        Frame1.Move -50, 0
ReadPakList
End Sub

Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu mnuFile
End If
End Sub

Private Sub mnuDelete_Click()
On Error Resume Next
    DeleteAll
End Sub

Private Sub mnuFileAdd_Click()
OpenFile Me
GetBinSize
SaveFileInPak
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileExtract_Click()
If frmMain.lstFiles.ListIndex = 0 Then Exit Sub
SaveFile Me
SimpleExtract
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "AppenderPak v" & App.Major & "." & App.Minor & "." & App.Revision & Chr(10) & Chr(10) & "Simple Pak creator with no compression." _
& Chr(10) & Chr(10) & "By:  Michael Heath", vbInformation, "About AppenderPak"
End Sub

Private Sub mnuHelpCredits_Click()
MsgBox "I would like to extend a special thanks to Robert Carter for his " _
& Chr(10) & "'PutFileInString' function and Tim Butler for helping me brain bash " _
& Chr(10) & "the code for fixing the funky offset.", vbInformation, "AppenderPak - Special Thanks"
End Sub
