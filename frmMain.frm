VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   " Batch VB6 Compiler"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8310
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin Project1.splitter splitter1 
      Height          =   45
      Left            =   0
      TabIndex        =   11
      Top             =   3600
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   79
      SplitterDirection=   1
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmMain.frx":058A
      Top             =   330
      Width           =   1215
   End
   Begin VB.Frame framCompile 
      Height          =   3195
      Left            =   1290
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   2985
      Begin VB.Frame frmSlide 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2715
         Left            =   60
         TabIndex        =   2
         Top             =   120
         Width           =   7815
         Begin Project1.chameleonButton cmdCancel 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   1380
            TabIndex        =   12
            Top             =   2160
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   926
            BTYPE           =   3
            TX              =   "Cancel"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   12632319
            BCOLO           =   64
            FCOL            =   0
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMain.frx":0593
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00FFFFFF&
            Height          =   2695
            Left            =   0
            ScaleHeight     =   2640
            ScaleWidth      =   705
            TabIndex        =   4
            Top             =   0
            Width           =   765
            Begin VB.PictureBox picArrow 
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BorderStyle     =   0  'None
               Height          =   1710
               Left            =   150
               Picture         =   "frmMain.frx":05AF
               ScaleHeight     =   1710
               ScaleWidth      =   405
               TabIndex        =   9
               Top             =   480
               Width           =   405
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   90
               Picture         =   "frmMain.frx":0BF7
               Top             =   0
               Width           =   480
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   150
               Picture         =   "frmMain.frx":0F2B
               Top             =   2220
               Width           =   480
            End
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   285
            Left            =   1380
            TabIndex        =   3
            Top             =   1800
            Width           =   6405
            _ExtentX        =   11298
            _ExtentY        =   503
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Current Project:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   8
            Top             =   30
            Width           =   1350
         End
         Begin VB.Label lblProject 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   7
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Current EXE File:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1470
            TabIndex        =   6
            Top             =   630
            Width           =   1350
         End
         Begin VB.Label lblEXE 
            AutoSize        =   -1  'True
            Caption         =   "-------"
            Height          =   195
            Left            =   1470
            TabIndex        =   5
            Top             =   840
            Width           =   420
         End
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7620
      Top             =   2670
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "BPR files (*.bpr)|*.BPR"
   End
   Begin MSComctlLib.ListView lstProj 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Project Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EXE Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Batch Profile"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Batch Profile..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Batch Profile..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save Batch Profile &As..."
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuCompile 
      Caption         =   "&Compile"
      Begin VB.Menu mnuCompileAll 
         Caption         =   "&Compile All"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCompileSingle 
         Caption         =   "Compile &Selected Project(s)"
      End
   End
   Begin VB.Menu mnuQSave 
      Caption         =   "&Quick Save"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EditResult As Integer
Dim SaveLocation As String
Dim mChanged As Boolean
Dim mCompiling As Boolean
Dim CancelLoop As Boolean

Private Sub cmdCancel_Click()
  
  'Set Cancel Variable to cancel current compilation loop
  CancelLoop = True
  
End Sub

Private Sub Form_Load()
  
  Dim Answer As Integer
  
  'Check for VB Path in Registry
  VBPath = GetSetting("BatchCompile", "Settings", "VBPath")
  
  'If it doesnt exist, do a scan.
  If VBPath = "" Then
    
    If Command$ <> "" Then
      
      'if a drive was passed, do not display message.
      Answer = vbYes
      
    Else
      
      Answer = MsgBox("NOTE: If VB6 is not installed on your C: Drive, This scan will fail." & vbCrLf & vbCrLf & _
                     "Please compile the EXE file, and pass the drive VB6 is installed on" & vbCrLf & _
                     "as the command$ FOR THE FIRST RUN ONLY. Once VB6 is found," & vbCrLf & _
                     "you can run this code anyway you'd like." & vbCrLf & vbCrLf & _
                     "Example (installed on S: Drive):" & vbCrLf & _
                     "c:\vbcode\batchcompiler.exe s" & vbCrLf & vbCrLf & "Continue to Scan C: Drive?", vbYesNo + vbQuestion, "Scanning for VB6.exe")
    
    End If
    
    'Continue with Scan on C:?
    If Answer = vbYes Then
      
      frmFindVB.Show
      Unload frmFindVB
      
      If VBPath = "" Then
        MsgBox "VB Path Not Found."
        End
      Else
        'Save VBPath to Registry
        SaveSetting "BatchCompile", "Settings", "VBPath", VBPath
        MsgBox "To add a *.VBP Project to your current Batch Profile," & vbCrLf & "simply drag the *.VBP file into the list.", vbInformation
      End If
    Else
      'Dont want to continue Scan on C:? go here.
      End
    End If
  End If
  
  txtLog = ""
  
  'Ensure form is visible
  Visible = True
  Do Until Visible
    DoEvents
  Loop
  
  'Load Last profile.
  Dim LastProfile As String
  LastProfile = GetSetting("BatchCompile", "Settings", "LastProfile")
  If LastProfile <> "" Then
    Call OpenProfile(LastProfile)
    CD.InitDir = GetDirectoryName(LastProfile)
  End If
  
End Sub

Private Sub Form_Resize()
  'Resizeing Controls and column headers

  
  lstProj.Move 0, 0, ScaleWidth, (ScaleHeight / 100) * 75
  txtLog.Move 0, lstProj.Height, ScaleWidth, (ScaleHeight / 100) * 25
  framCompile.Move 0, -60, ScaleWidth, lstProj.Height
  frmSlide.Move 60, 240, framCompile.Width - 120
  splitter1.Width = ScaleWidth
  splitter1.Top = txtLog.Top - splitter1.Height
  
  Dim X As Integer
  For X = 1 To lstProj.ColumnHeaders.Count
    lstProj.ColumnHeaders(X).Width = (lstProj.Width - 120) / lstProj.ColumnHeaders.Count
  Next
  
  If Height < 5145 Then
    LockWindow Me.hWnd, True
    Height = 5145
    LockWindow Me.hWnd, False
  End If
  
  
  If mCompiling Then
    Set splitter1.Control_Top_Or_Left = framCompile
  Else
    Set splitter1.Control_Top_Or_Left = lstProj
  End If
  Set splitter1.Control_Bottom_Or_Right = txtLog
End Sub

Private Sub lstProj_DblClick()
  
  'Load a new Project Editor Form
  Dim Another As New frmEditProject
  
  'Set the mProj to the project arrays index
  Another.mProj = Val(Replace(LCase$(lstProj.SelectedItem.Key), "key:", ""))
  
  Another.Show 1
  
  If EditResult = vbOK Then
    DoChanged (True)
    PopulateProjects
  End If
  
End Sub

Private Sub lstProj_KeyUp(KeyCode As Integer, Shift As Integer)
  Dim i As Integer
  
  'To Delete a project
  If KeyCode = 46 Then 'Delete was Pressed
    Dim Answer As Integer
    i = Val(Replace(LCase$(lstProj.SelectedItem.Key), "key:", ""))
    Answer = MsgBox("Are you sure you want to delete this project from the profile?" & vbCrLf & vbCrLf & lstProj.SelectedItem.Text, vbYesNo + vbQuestion, "Delete Project From Profile?")
    If Answer = vbYes Then
      Dim X As Integer
      
      'move entries down to fill deleted spot
      For X = i To UBound(mProjects) - 1
        mProjects(X).ExeFullPath = mProjects(X + 1).ExeFullPath
        mProjects(X).EXEName = mProjects(X + 1).EXEName
        mProjects(X).ProjectFullPath = mProjects(X + 1).ProjectFullPath
        mProjects(X).ProjectName = mProjects(X + 1).ProjectName
      Next
      
      'Clip off the last array slot
      ReDim Preserve mProjects(UBound(mProjects) - 1)
      
      'Refresh List
      PopulateProjects
      
      'Set CHANGED to true
      DoChanged (True)
    End If
    
  ElseIf KeyCode = 116 Then 'F5
  
      'Refresh List
      PopulateProjects
  
  End If
End Sub

Private Sub lstProj_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, y As Single)
  If Data.GetFormat(vbCFFiles) Then
    Dim tFile
    'Receives Dragged/Dropped Files
    For Each tFile In Data.Files
      'Make sure its a VBP file
      If Right(LCase(tFile), 4) = ".vbp" Then
        Call DoChanged(True)
        Call AddProject(tFile)
      End If
    Next
    Call PopulateProjects
  End If
End Sub

Sub AddProject(vFile)
  On Error GoTo Err
  Dim pFile As String
  pFile = vFile
  
  
  Dim Readln As String
  Dim Num As Integer
  Dim Path32 As String
  Dim ExeName32 As String
  Num = FreeFile()
  
  'Retreive Compile information from VBP file
  Open vFile For Input As #Num
    Do While Not EOF(Num)
      Line Input #Num, Readln
      If Left(Readln, 9) = "ExeName32" Then
        ExeName32 = Replace(Readln, "ExeName32=" & Chr(34), "")
        ExeName32 = Left$(ExeName32, Len(ExeName32) - 1)
      ElseIf Left(Readln, 6) = "Path32" Then
        Path32 = Replace(Readln, "Path32=" & Chr(34), "")
        Path32 = Left$(Path32, Len(Path32) - 1)
      End If
    Loop
  Close #Num
  
  'Project hasnt been compiled yet?
  If ExeName32 = "" Then
    MsgBox "Please compile this project manually in VB First."
    Exit Sub
  End If
  
  'Get Realtive Path to the project file
  Path32 = GetRelativePath(Path32, Left$(pFile, Len(pFile) - (Len(GetFileName(pFile))) - 1))
    
  Dim EXEName As String
  EXEName = Path32 & "\" & ExeName32
  EXEName = Replace(EXEName, "\\", "\")
  
  'Check to see if the project already exists in the list
  If UBound(mProjects()) > 0 Then
    Dim X As Integer
    For X = 0 To UBound(mProjects)
      If LCase$(mProjects(X).ProjectFullPath) = LCase(pFile) Then
        Exit Sub
      End If
    Next
  End If
  
  GoTo 10
  
Err:
  
  'This indicates an empty array
  ReDim Preserve mProjects(0)
  
  GoTo 20
  
10:
  
  If lstProj.ListItems.Count = 0 Then
    ReDim Preserve mProjects(0)
  Else
    ReDim Preserve mProjects(UBound(mProjects) + 1)
  End If
  
20:

  mProjects(UBound(mProjects)).ProjectFullPath = pFile
  mProjects(UBound(mProjects)).ProjectName = GetFileName(mProjects(UBound(mProjects)).ProjectFullPath)
  
  mProjects(UBound(mProjects)).ExeFullPath = EXEName
  mProjects(UBound(mProjects)).EXEName = LCase$(GetFileName(mProjects(UBound(mProjects)).ExeFullPath))
  
End Sub

Function GetRelativePath(findPath As String, startPath As String) As String
  
  Dim L As Integer
  Dim i As Integer
  Dim Backs As Integer
  
  'Find out how many BackDirs (..\) there are
  L = Len(findPath)
  findPath = Replace(findPath, "..\", "")
  Backs = (L - Len(findPath)) / 3
  
  'Back up BACKS BackDirs
  For i = 1 To Backs
    If i = 1 Then
      L = InStrRev(startPath, "\")
    Else
      L = InStrRev(startPath, "\", L - 1)
    End If
    startPath = Left(startPath, L - 1)
  Next
  
  GetRelativePath = startPath & "\" & findPath
  
End Function

Sub Dostatus(vProjName As String, vEXEName As String, vPercent As Single)
  
  'If not compiling, this function shouldnt be called
  If Not (framCompile.Visible) Then Exit Sub
  
  'Display current compile information
  lblProject = vProjName
  lblEXE = vEXEName
  ProgressBar1.Value = Int(vPercent)
End Sub

Sub AllControlsEnabled(vEnabled As Boolean)
  mnuFile.Enabled = vEnabled
  mnuCompile.Enabled = vEnabled
  lstProj.Visible = vEnabled
  framCompile.Visible = Not (vEnabled)
  If vEnabled = False Then
    mnuQSave.Enabled = False
  Else
    mnuQSave.Enabled = mChanged
  End If
End Sub

Sub CompileFunc(vSelectedOnly As Boolean)
    
  mCompiling = True
  
  Dim X As Integer
  Dim i As Integer
  Dim CMD As String
  Dim Cnt As Integer
  Dim Cnt2 As Integer
  Dim Total As Integer
  'Disable Controls
  
  AllControlsEnabled (False)
  
  Set splitter1.Control_Top_Or_Left = framCompile
  Set splitter1.Control_Bottom_Or_Right = txtLog
  
  DoEvents

  'Count Selected Items, if needed.
  If vSelectedOnly Then
    For X = 1 To lstProj.ListItems.Count
      If lstProj.ListItems(X).Selected Then Cnt = Cnt + 1
    Next
  End If

  Total = IIf(vSelectedOnly, Cnt, lstProj.ListItems.Count)

  'Go Through List
  For X = 1 To lstProj.ListItems.Count
    
    'Check for Cancel Being Pressed
    If CancelLoop Then
      
      'Add the Alert to the Log
      Call AddAlertToLog(txtLog, "Operation Aborted by User")
      
      'Reset Cancel Variable
      CancelLoop = False
      
      'Stop the Loop
      Exit For
      
    End If
    
    If (vSelectedOnly And lstProj.ListItems(X).Selected) Or Not (vSelectedOnly) Then
      
      'Increment cnt2
      Cnt2 = Cnt2 + 1
    
      'Show 'Compiling' Arrow
      picArrow.Visible = True
      
      'Let Process Finish
      DoEvents
      
      'Find Array's Index Value
      i = Val(Replace(LCase$(lstProj.ListItems(X).Key), "key:", ""))
      
      'Build the command to shell
      CMD = VBPath & " /make " & Chr(34) & mProjects(i).ProjectFullPath & Chr(34) & " " & Chr(34) & mProjects(i).ExeFullPath & Chr(34)
      
      'Update Status Frame
      Call Dostatus(mProjects(i).ProjectFullPath, mProjects(i).ExeFullPath, (Cnt2 / Total) * 100)
      
      'Shell the command
      Call Shell(CMD)
      
      'Add Compile to Log
      Call CompileToLog(txtLog, mProjects(i).ProjectFullPath, mProjects(i).ExeFullPath, Cnt2, IIf(vSelectedOnly, Cnt, lstProj.ListItems.Count))
      
      'Hide 'Compiling' Arrow
      picArrow.Visible = False
      
      'Pause - this is for effect so we actuall see the arrow when
      'Compiling relatively small projects
      Pause (0.1)
        
    End If
    
    'Let Process Finish
    DoEvents
    
  Next

  Set splitter1.Control_Top_Or_Left = lstProj
  Set splitter1.Control_Bottom_Or_Right = txtLog

  Call AddAlertToLog(txtLog, "Finished Compiling " & Cnt2 & " Projects")

  'Re-Enable Controls
  AllControlsEnabled (True)
  mCompiling = False
End Sub

Private Sub mnuCompileSingle_Click()
  
  'Compile Select Projects
  Call CompileFunc(True)

End Sub

Sub mnuCompileAll_Click()

  'Compile All Projects
  Call CompileFunc(False)

End Sub

Private Sub mnuExit_Click()

  'Unload the form, causing the application to end
  Unload Me

End Sub

Private Sub mnuNew_Click()
  
  'Reset the save location, so we are prompted next time.
  SaveLocation = ""
  
  'Clear List
  lstProj.ListItems.Clear
  
  'Reset Caption
  DoCaption ("")
  
  'Clear Projects Array
  ReDim mProjects(0)
  
  'Set Changed=False
  Call DoChanged(False)
  
End Sub

Private Sub mnuOpen_Click()

  'If Error is detected, Exit the sub
  On Error GoTo Err
  
  'If cancel is pressed, generate the error to exit sub.
  CD.CancelError = True
  
  'Show Open Dialog
  CD.ShowOpen
  
  'Make call to open then file
  Call OpenProfile(CD.FileName)
  
Err:
End Sub

Sub OpenProfile(vFile As String)
  On Error GoTo Err
  
  'Just in case a blank file is passed, exit sub
  If Dir(vFile) = "" Then Exit Sub
  
  Dim Count As Integer
  Dim fNum As Integer
  
  'Find next open file number to use
  fNum = FreeFile()
  
  'Simple File read to populate Data
  Open vFile For Input As #fNum
    Do While Not EOF(fNum)
      
      'This will handle the array
      If Count = 0 Then
        
        'if count is 0, create a new, blank array
        ReDim mProjects(0)
        
      Else
        
        'If not, Preserve the exsisting array, but add a new slot
        ReDim Preserve mProjects(Count)
        
      End If
      
      'Read the Project Path, and assign it to the Project Variable
      Line Input #1, mProjects(Count).ProjectFullPath
      
      'Set the Project Name, using the Project Path
      mProjects(Count).ProjectName = GetFileName(mProjects(Count).ProjectFullPath)
      
      'Read the Exe Path, and assign it to the Project Variable
      Line Input #1, mProjects(Count).ExeFullPath
      
      'Set the Exe Name, using the Exe Path
      mProjects(Count).EXEName = GetFileName(mProjects(Count).ExeFullPath)
      
      'Increment the counter
      Count = Count + 1
      
    Loop
    
  'close the file
  Close #fNum
  
  'Using the Projects Array, Populate the list view
  Call PopulateProjects
  
  'Update Caption
  DoCaption Left(GetFileName(vFile), Len(GetFileName(vFile)) - 4)
  
  'Save Settings
  SaveLocation = vFile
  SaveSetting "BatchCompile", "Settings", "LastProfile", vFile
  
  'Since this is a newly opened project, nothing has been changed.
  Call DoChanged(False)
  
  Exit Sub
  
Err:
  MsgBox Err.Description
End Sub

Sub SaveProfile(Optional vSaveAs As Boolean)
  On Error GoTo Err
  
  'If SaveLocation is blank, or 'Save As' was clicked, show the dialog
  If SaveLocation = "" Or vSaveAs Then
    
    'Trigger Error to Exit sub if Cancel was Pressed
    CD.CancelError = True
    
    'Show save Dialog box
    CD.ShowSave
    
    'Set the current Save location to the file selected
    SaveLocation = CD.FileName
  End If
  
  Dim fNum As String
  
  'Find Next Open File Number
  fNum = FreeFile()
  
  'Check for over-write
  Dim Answer As Integer
  If vSaveAs And Dir(SaveLocation) <> "" Then
    
    'If trying to save over an existing file, prompt to confirm.
    Answer = MsgBox("The file already exists. OK to over-write?", vbYesNo + vbQuestion, "Over-Write?")
    
  ElseIf Not vSaveAs Then
    
    'if the user just presses 'Save' and we already know SaveLocation, just save
    'over existing file.
    Answer = vbYes
    
  End If
  
  'Should we delete the existing file?
  If Answer = vbYes Then
    
    'If we want to delete the file, lets do it.
    If Dir(SaveLocation) <> "" Then Kill SaveLocation
    
  Else
  
    'if not, we can not save.
    Exit Sub
    
  End If
   
  'Simple File Write to save settings
  Dim X As Integer
  
  'Open Output file
  Open SaveLocation For Output As #fNum
    
    'Loop through projects
    For X = 0 To UBound(mProjects)
    
      'Print Data to file
      Print #fNum, mProjects(X).ProjectFullPath
      Print #fNum, mProjects(X).ExeFullPath
      
    Next
    
  'Close File
  Close #fNum
  
  'Update Caption
  DoCaption Left(GetFileName(SaveLocation), Len(GetFileName(SaveLocation)) - 4)
  
  'SaveSettings
  SaveSetting "BatchCompile", "Settings", "LastProfile", SaveLocation
  
  'Since we've just saved, there are no changes to this project
  Call DoChanged(False)
  
Err:
End Sub

Sub DoChanged(vChanged As Boolean)
  'Set mChanged Varialabl
  mChanged = vChanged
  
  'Set Control(s) appropriately
  mnuQSave.Enabled = vChanged
End Sub

Sub PopulateProjects()
  Dim X As Integer
  Dim LI As ListItem
  
  LockWindow lstProj.hWnd, True
  
  'Delete all Current Items
  lstProj.ListItems.Clear
  
  'populate listview with array data
  For X = 0 To UBound(mProjects())
    
    'Set List Item Text (Cell 1)
    Set LI = lstProj.ListItems.Add(X + 1, "Key:" & X, mProjects(X).ProjectName)
    
    'Set List Item's SubItem Text (Cell 2)
    LI.SubItems(1) = mProjects(X).EXEName
    
  Next
  
  LockWindow lstProj.hWnd, False
  
End Sub

Private Sub mnuQSave_Click()
  
  'Call Save Routine
  SaveProfile
  
End Sub

Private Sub mnuSave_Click()
  
  'Call Save Routine
  SaveProfile
  
End Sub

Private Sub mnuSaveAs_Click()
  
  'Call Save As Routine
  SaveProfile (True)
  
End Sub

Sub DoCaption(vText As String)
  
  'Set Application's Caption
  If vText = "" Then
    
    'Case to reset text to default
    Caption = " Batch VB6 Compiler"
  
  Else
    
    'Case to display information
    Caption = " Batch VB6 Compiler - " & vText
  
  End If

End Sub
