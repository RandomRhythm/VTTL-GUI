VERSION 5.00
Begin VB.Form frmLaunch 
   Caption         =   "VTTL Launch Arguments"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLaunch.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8505
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnStatus 
      Caption         =   "Refresh status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   22
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   960
      TabIndex        =   14
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   16000
      Left            =   2640
      Top             =   1200
   End
   Begin VB.CommandButton btnBrowse 
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start VTTL Lookup Script"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   15
      Top             =   5640
      Width           =   3495
   End
   Begin VB.TextBox txtImport 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   11
      ToolTipText     =   "Path to CSV file that will be merged with VTTL CSV output"
      Top             =   3600
      Width           =   4215
   End
   Begin VB.CheckBox chkTIA 
      Caption         =   "Disable TIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   8
      ToolTipText     =   "API key is required - threatintelligenceaggregator.org"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox chkThreatGRID 
      Caption         =   "Disable ThreatGRID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CheckBox chkET 
      Caption         =   "Disable ET Intelligence"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CheckBox chkAlien 
      Caption         =   "Disable AlienVault OTX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.OptionButton optImport 
      Caption         =   "Tabbed (EnCase)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   2520
      TabIndex        =   13
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox chkCBr 
      Caption         =   "Disable Cb Response API "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   480
      Width           =   2535
   End
   Begin VB.CheckBox chkImport 
      Caption         =   "Import fields from file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Tag             =   "1"
      ToolTipText     =   "Merge with sigcheck or other supported CSV. Only works with hash lookups"
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkMalshare 
      Caption         =   "Disable Malshare"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CheckBox chkThreatCrowd 
      Caption         =   "Disable ThreatCrowd"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "ThreatCrowd only do 4 lookups per minute or you may get banned"
      Top             =   1560
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox chkExcel 
      Caption         =   "Use Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Use Microsoft Excel instead of CSV"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox chkSilent 
      Caption         =   "Slient Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Disable message prompts"
      Top             =   840
      Width           =   1215
   End
   Begin VB.OptionButton optImport 
      Caption         =   "CSV (sigcheck/autorunsc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Value           =   -1  'True
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   6495
      Begin VB.Label lblFilePath 
         Caption         =   "File path:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   20
      Top             =   6600
      Width           =   6495
      Begin VB.Label lblStatus 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   4455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Include the following text in the output CSV name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   19
      Top             =   4440
      Width           =   5055
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   555
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Dim debugPath 'used to load debug path for use with status check

Private Sub btnBrowse_Click()
frmLaunch.txtImport.Text = MyDialogOpenCode
End Sub

Function ProcessQuery()
Dim killProcesses: killProcesses = False
Dim promptKill: promptKill = True
strComputer = "."
strProcID = GetCurrentProcessId()
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_Process" & _
           " WHERE ParentProcessId = '" & strProcID & "'", , 48)
For Each objItem In colItems
    If InStr(LCase(objItem.CommandLine), "vttl.vbs") > 0 And (LCase(objItem.Name) = "wscript.exe" Or LCase(objItem.Name) = "cscript.exe") Then
        If promptKill = True Then
            promptAnswer = MsgBox("A VTTL script instance is already runnning. Only one instance can run at a time. Do you want to stop the running instance to start a new one?", vbYesNo)
            If promptAnswer = vbYes Then
                'kill processes
                killProcesses = True
            End If
            promptKill = False
        End If
        If killProcesses = True Then
            killProc (objItem.ProcessId)
        End If
            


    End If
Next
If promptKill = True Then 'never saw VTTL running
    ProcessQuery = True
Else 'saw VTTL running. Return if we killed it or not
    ProcessQuery = killProcesses
End If
End Function
Sub killProc(strProcID)
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
& "{impersonationLevel=impersonate}!\\" _
& strComputer & "\root\cimv2")

Set colProcess = objWMIService.ExecQuery _
("Select * from Win32_Process Where ProcessId = " & strProcID)
For Each objProcess In colProcess
    objProcess.Terminate
Next
End Sub
Private Sub btnStart_Click()
check_type 'change lookup mode
If Timer1.Enabled = True Then Exit Sub
If Timer1.Enabled = False Then Timer1.Enabled = True
Set fso = CreateObject("Scripting.FileSystemObject")
Dim args
args = ""

If Form1.boolInitialInstall = False Then
    If frmLaunch.chkSilent.Value = 1 Then args = args & " /s"
    If frmLaunch.chkMalshare.Value = 1 Then args = args + " /dms"
    If frmLaunch.chkCbR.Value = 1 Then args = args + " /dcb"
    If frmLaunch.chkAlien.Value = 1 Then args = args + " /dav"
    If frmLaunch.chkET.Value = 1 Then args = args + " /det"
    If frmLaunch.chkThreatGRID.Value = 1 Then args = args + " /dtg"
    If frmLaunch.chkTIA.Value = 1 Then args = args + " /dtia"
    If frmLaunch.chkExcel.Value = 1 Then args = args + " /e"
    If frmLaunch.chkImport.Value = 1 And frmLaunch.txtImport.Text = "" Then
        If getStatusMessage = "Finished" Then
            theAnswer = vbYes
        Else
            theAnswer = MsgBox("Do you want to specify an input source to combine with VTTL output?", vbYesNo, "VTTL - " & App.Path)
        End If
         If theAnswer = vbYes Then
             frmLaunch.txtImport.SetFocus
             Exit Sub
         Else
             frmLaunch.chkImport.Value = 0
         End If
    End If
    If frmLaunch.chkImport.Value = 1 Then
        If Form1.optHash(0).Value = True Then 'hash lookups
            If optImport(0).Value = 1 Then
                args = args + " /a"
            Else
                args = args + " /g"
            End If
        ElseIf optImport(0).Value = 0 Then 'domain IP lookups
            args = args + " /p"
        End If
        If txtImport.Text <> "" Then
            If fso.FileExists(txtImport.Text) Then
                args = args & " " & Chr(34) & txtImport.Text & Chr(34)
            End If
        End If
    End If
    'MsgBox args
End If
If frmLaunch.chkThreatCrowd.Value = 1 Then args = args + " /dtc"
If txtFileName.Enabled = True And txtFileName.Text <> "" Then args = args & " /n " & txtFileName.Text

strLaunch = "wscript.exe "

If fso.FileExists("C:\Windows\SysWOW64\wscript.exe") Then
    strLaunch = "c:\windows\sysnative\wscript.exe "
End If

If ProcessQuery = False Then
    Timer1.Enabled = False 'disable timer so user can hit button again
    Exit Sub
End If
If fso.FileExists(App.Path & "\vttl.vbs") = False Then
    MsgBox "vttl.vbs does not exist: " & App.Path & "\vttl.vbs"
Else
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.CurrentDirectory = App.Path
    WshShell.Run strLaunch & Chr(34) & App.Path & "\vttl.vbs" & Chr(34) & args, 0
    Set WshShell = Nothing
    frmLaunch.lblStatus.Caption = "Starting"
    frmLaunch.Height = 8970
End If


Form1.boolInitialInstall = False
IniSection = "main"
strIniKey = "InitialInstall"
Form1.WriteINIvalue IniSection, strIniKey, False
End Sub



Private Sub btnStatus_Click()

    frmLaunch.lblStatus.Caption = getStatusMessage
    If frmLaunch.lblStatus.Caption = "Status unknown" Then
        MsgBox "Problem locating status file:" & statusFile
    End If
End Sub

Function getStatusMessage()
statusFile = debugPath & "\status.txt"
statusMessage = ""
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(statusFile) = False Then
    '
    frmLaunch.lblStatus.Caption = "Status unknown"
    Exit Function
End If

    Set objTextFile = objFSO.OpenTextFile _
    (statusFile, ForReading)
    If objTextFile.AtEndOfStream = False Then
        statusMessage = objTextFile.ReadAll
    End If
    objTextFile.Close
    getStatusMessage = statusMessage
End Function

Private Sub chkImport_Click()
If chkImport.Value = 0 Then
    optImport(0).Enabled = False
    optImport(1).Enabled = False
    txtImport.Enabled = False
    btnBrowse.Enabled = False
    chkImport.Tag = 0
Else
    optImport(0).Enabled = True
    optImport(1).Enabled = True
    txtImport.Enabled = True
    btnBrowse.Enabled = True
    chkImport.Tag = 1
End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub EnableCheckControl(strDatPath, chkControl, from1ChkControl)
Set fsoECC = CreateObject("Scripting.FileSystemObject")
If fsoECC.FileExists(strDatPath) = True Then
    chkControl.Enabled = True
    If fsoECC.FileExists(Replace(LCase(strDatPath), ".dat", ".disable")) = True Then
        chkControl.Value = 1
        chkControl.Enabled = False 'if a disable file was created they decided to not use this feature
    Else
        chkControl.Value = 0
        chkControl.Enabled = True
    End If
ElseIf from1ChkControl.Value = 1 Then
    chkControl.Value = 0 'Need to ensure we prompt the user to create the dat file
Else
    'chkControl.Enabled = False 'Allow them to enable so script will prompt for API key to create dat
    chkControl.Value = 1
End If
End Sub

Public Sub SetFormEnables(boolStartup)
If boolStartup = False And frmLaunch.Visible = False Then Exit Sub
If Form1.chkAlienNIDS.Value = 1 Then frmLaunch.chkAlien = 0
If Form1.chkCbR.Value = 1 Then frmLaunch.chkCbR = 0
frmLaunch.Caption = "VTTL Launch Arguments - " & App.Path
If Form1.optHash(1) = 0 Then
    lblTitle.Caption = "Hash Lookup Mode"
    
    frmLaunch.chkImport.Enabled = True
    'optImport(0).Enabled = True
    'optImport(1).Enabled = True
    optImport(0).Visible = True
    optImport(1).Caption = "CSV (sigcheck/autorunsc)"
    
    'txtImport.Enabled = True
    'frmLaunch.btnBrowse.Enabled = True
    chkMalshare.Enabled = True
    chkCbR.Enabled = True
    chkTIA.Enabled = True
    'lblFilePath.Enabled = True
    If chkImport.Tag = 1 Then frmLaunch.chkImport.Value = 1
    EnableCheckControl App.Path & "\cb.dat", frmLaunch.chkCbR, Form1.chkCbR
    EnableCheckControl App.Path & "\tia.dat", frmLaunch.chkTIA, Form1.chkTIA
    EnableCheckControl App.Path & "\malshare.dat", frmLaunch.chkMalshare, Form1.chkMalshare
Else
    lblTitle.Caption = "IP/Domain Lookup Mode"
    'frmLaunch.chkImport.Enabled = False
    importSetting = frmLaunch.chkImport.Tag
    frmLaunch.chkImport.Value = 0
    frmLaunch.chkImport.Tag = importSetting
    txtImport.Text = ""
    optImport(0).Enabled = False
    optImport(1).Caption = "Custom CSV - Prevalence and Sibling Count"
    optImport(1).Enabled = False
    optImport(0).Visible = False
    txtImport.Enabled = False
    frmLaunch.btnBrowse.Enabled = False
    
    'lblFilePath.Enabled = False
    chkMalshare.Enabled = False
    chkCbR.Enabled = False
    chkTIA.Enabled = False
End If
Set fsoFL = CreateObject("Scripting.FileSystemObject")



EnableCheckControl App.Path & "\et.dat", frmLaunch.chkET, Form1.chkET
EnableCheckControl App.Path & "\tg.dat", frmLaunch.chkThreatGRID, Form1.chkThreatGRID


intDelay = Form1.txtTimeDelay
If IsNumeric(intDelay) Then ' VirusTotal sleep disabled and db caching enabled
    If intDelay < 10000 Or (Form1.ChkSleep.Value = 0 And Form1.chkReadCache = 0) Then
        chkThreatCrowd.Value = 1
        chkThreatCrowd.Enabled = False
        chkMalshare.Value = 1
        chkMalshare.Enabled = False
    Else
        chkThreatCrowd.Enabled = True
        If lblTitle.Caption <> "IP/Domain Lookup Mode" Then chkMalshare.Enabled = True
        chkMalshare.Enabled = True
    End If
End If




EnableCheckControl App.Path & "\av.dat", frmLaunch.chkAlien, Form1.chkAlienVault

If Form1.chkThreatCrowd.Value = 1 Then
    frmLaunch.chkThreatCrowd.Value = 0
End If

If Form1.chkMalshare.Value = 0 Then
    frmLaunch.chkMalshare.Value = 1
End If
End Sub

Function MyDialogOpenCode()
Dim objDialog
Dim IsW2k As Boolean
Dim intReturn
On Error Resume Next
Set objDialog = CreateObject("SAFRCFileDlg.FileOpen")
If Err.Number <> 0 Then
  IsW2k = True
Else
  IsW2k = False
End If
On Error GoTo 0

If IsW2k = False Then

    
    
      intReturn = objDialog.OpenFileOpenDlg
    MyDialogOpenCode = objDialog.FileName
  
    Set objDialog = Nothing
Else
    
    'CommonDialog1.ShowOpen
    'MyDialogOpenCode = CommonDialog1.FileName
    'MyDialogSaveCode = InputBox("Please type the file path where you would like to open the file")
    Dim c As New cCommonDialog
  
    c.ShowOpen
    MyDialogOpenCode = c.FileName
  
    Set c = Nothing
    
End If





End Function

Private Sub Form_Load()
frmLaunch.Height = 6945
check_type
strIniPath = App.Path & "\vttl.ini"
strIniSection = "Debug"
strIniKey = "path"
debugPath = Form1.readINIvalue(strIniSection, strIniPath, strIniKey) '
If debugPath = "" Then
    debugPath = App.Path & "\Debug\Operations"
End If
If lblTitle.Caption = "Hash Lookup Mode" Then
    optImport(1).ToolTipText = "CSV tool output to merge with VTTL CSV output"
    
Else
    optImport(1).ToolTipText = "CSV tool output to merge with VTTL CSV output. Header row must use " & Chr(34) & "Prevalence" & Chr(34) & " and/or " & Chr(34) & "Sibling Count" & Chr(34) & " as the header values"
End If
End Sub
Sub check_type()

If Form1.chkExcel.Value = 1 Then
    txtFileName.Enabled = False
    txtFileName.ToolTipText = "Disabled when using Excel"
Else
    txtFileName.Enabled = True
End If
End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = False
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    frmLaunch.btnStart.SetFocus
    btnStart_Click
End If
End Sub
