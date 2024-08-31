VERSION 5.00
Begin VB.Form frmFeed 
   Caption         =   "Feed Configuration"
   ClientHeight    =   8670
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11670
   Icon            =   "FrmFeed.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8670
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescColumn 
      Enabled         =   0   'False
      Height          =   495
      Left            =   9720
      TabIndex        =   10
      ToolTipText     =   "Starting at zero, what column number in the CSV contains the IOC description"
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtIntelColumn 
      Enabled         =   0   'False
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      ToolTipText     =   "Starting at zero, what column number in the CSV contains the IOC"
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox chkCSV 
      Caption         =   "Feed is CSV"
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      ToolTipText     =   "Set this for tab separated as well"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox ComboAgeFilter 
      Height          =   315
      Left            =   10080
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Timer timerFlashy 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3120
      Top             =   1920
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add New Feed"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   10935
   End
   Begin VB.CommandButton btnUpdateEntry 
      Caption         =   "Update The Selected Feed Entry With Changes"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   6480
      Width           =   10935
   End
   Begin VB.CommandButton btnStatic 
      Caption         =   "Browse Static Inteligence"
      Height          =   495
      Left            =   6000
      TabIndex        =   27
      ToolTipText     =   "Folder containing text files listing intelligence that are not updated automatically"
      Top             =   7800
      Width           =   5295
   End
   Begin VB.TextBox txtAgeLimit 
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      ToolTipText     =   "Default is blank. Only use if you want to be able to target the feed based on how far back the feed goes via the age filter"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CheckBox ChkIgnoreSSL 
      Caption         =   "Ignore SSL Errors"
      Height          =   495
      Left            =   6600
      TabIndex        =   3
      ToolTipText     =   "Ignore SSL errors when downloading the feed"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox txtRefresh 
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      ToolTipText     =   "Hours to wait before downloading an updated copy of the feed"
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtCacheLocation 
      Height          =   495
      Left            =   360
      TabIndex        =   6
      ToolTipText     =   "File name where to store the feed. Please use file extension .txt or .csv"
      Top             =   2520
      Width           =   3975
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      ToolTipText     =   "Enable/disables feed. Disabling will prevent feed being used even when category is enabled."
      Top             =   1440
      Width           =   975
   End
   Begin VB.ListBox lstFeed 
      Height          =   2400
      Left            =   360
      TabIndex        =   12
      Top             =   3960
      Width           =   10935
   End
   Begin VB.CommandButton btnResetFeed 
      Caption         =   "Reset Feeds to Default"
      Height          =   495
      Left            =   360
      TabIndex        =   29
      Top             =   7800
      Width           =   5535
   End
   Begin VB.TextBox txtURL 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   10935
   End
   Begin VB.TextBox txtFeedString 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Case sensitve text that must be present to accept the feed download as successful"
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ComboBox ComboCategory 
      Height          =   315
      Left            =   4800
      TabIndex        =   2
      Text            =   "Combo1"
      ToolTipText     =   "Category of feed allows disable/enable of many feeds based on category"
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox chk_Static 
      Caption         =   "Static Intel"
      Height          =   495
      Left            =   7920
      TabIndex        =   19
      ToolTipText     =   "Use static intelligence"
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox chkWatchIntelURL 
      Caption         =   "Watch Intel URLs"
      Height          =   495
      Left            =   6120
      TabIndex        =   18
      ToolTipText     =   "Add URLs from intel feeds to watchlist"
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CheckBox chk_Multi 
      Caption         =   "Multi Feed"
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox chk_proxy 
      Caption         =   "Proxy Feed"
      Height          =   495
      Left            =   4800
      TabIndex        =   17
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CheckBox chk_Attacker 
      Caption         =   "Attacker Feeds"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CheckBox chk_Malware 
      Caption         =   "Malware Feeds"
      Height          =   495
      Left            =   1920
      TabIndex        =   15
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label lblcolumn 
      Caption         =   "Intel Column"
      Height          =   495
      Left            =   8280
      TabIndex        =   30
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblDesc 
      Caption         =   "Description Column"
      Height          =   495
      Left            =   9720
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Age Filter:"
      Height          =   615
      Left            =   9240
      TabIndex        =   28
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Age Limit - Days"
      Height          =   495
      Left            =   6600
      TabIndex        =   26
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Refresh Age - Hours"
      Height          =   495
      Left            =   4800
      TabIndex        =   25
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Cache Location"
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Category 
      Caption         =   "Category"
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Feed String Check"
      Height          =   495
      Left            =   360
      TabIndex        =   22
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Feed URL"
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmFeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'upgrade process
'copy exe to target directory
'launch copied exe passing argument of file path to recently download repo
'ini is replaced and then configuration is saved.
'script is replaced

Dim boolSaveFeedChange As Boolean

Private Sub btnAdd_Click()
If btnAdd.Caption = "Add New Feed" Then
    btnAdd.Caption = "Cancel Feed Add"
    If FrmFeed.lstFeed.ListIndex <> -1 Then 'we did not just added this in a new  window without selecting in the list
        txtCacheLocation.Text = ""
    End If
    FrmFeed.lstFeed.ListIndex = -1
    lstFeed.Enabled = False
    btnUpdateEntry.Caption = "Add New Feed With Current Values"
    If txtRefresh.Text = "" Then txtRefresh.Text = 24
    FrmFeed.chkEnabled.Value = 1

Else
    btnAdd.Caption = "Add New Feed"
    lstFeed.Enabled = True
    btnUpdateEntry.Caption = "Update The Selected Feed Entry With Changes"
End If
End Sub

Private Sub btnResetFeed_Click()
intAnswer = MsgBox("Are you sure you want to reset? You will loose all custom feed configurations.", vbYesNo, "Reset Feeds")
If intAnswer = vbNo Then Exit Sub
chk_Attacker.Value = 1
chk_Malware.Value = 1
chk_Multi.Value = 1
chk_proxy.Value = 1
chk_Static.Value = 1
chkWatchIntelURL.Value = 0

'ComboCategory.AddItem "proxy"
'ComboCategory.AddItem "malware"
'ComboCategory.AddItem "attacker"
'ComboCategory.AddItem "multi"




        txtURL.Text = ""
        txtCacheLocation = ""
        txtFeedString.Text = ""
        
        chkEnabled.Value = 0
        txtRefresh.Text = ""
        
        ChkIgnoreSSL.Value = 0
        txtAgeLimit = ""
        ComboCategory = ""

feedRestore "feedlist.default", "feedlist.dat"
feedRestore "csvfeed.default", "csvfeed.dat"

lstFeed.Clear
ReadFeedlist "", "", False, True, False 'use "" parameter to load dict
Form1.dictCSVFeed.RemoveAll
Form1.LoadCustomValDict App.Path & "\config\csvfeed.dat", dictCSVFeed
End Sub

Sub feedRestore(feedBackup, feedRestore)

Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(App.Path & "\config\" & feedBackup) Then
    If objFSO.FileExists(App.Path & "\config\" & feedRestore) Then objFSO.DeleteFile (App.Path & "\config\" & feedRestore)
    objFSO.CopyFile App.Path & "\config\" & feedBackup, App.Path & "\config\" & feedRestore
Else
    MsgBox "Missing file:" & App.Path & "\config\" & feedBackup
End If

End Sub

Private Sub btnStatic_Click()
strIniPath = App.Path & "\vttl.ini"
strIniSection = "main"
strIniKey = "StaticIntelPath"
staticIntelPath = Form1.readINIvalue(strIniSection, strIniPath, strIniKey) 'Restrict timeframe of event instances


Set objFSO = CreateObject("Scripting.FileSystemObject")
If staticIntelPath = "" Or staticIntelPath = " " Then

    staticIntelPath = "\static"
  End If

  If objFSO.FolderExists(staticIntelPath) Then
    'all set
  ElseIf objFSO.FolderExists(App.Path & "\" & staticIntelPath) Then
    staticIntelPath = App.Path & "\" & staticIntelPath
  Else
    objFSO.CreateFolder (App.Path & "\" & staticIntelPath)
    staticIntelPath = App.Path & "\" & staticIntelPath
  End If
staticIntelPath = Replace(staticIntelPath, "\\", "\")
 Set WshShell = CreateObject("WScript.Shell")
 WshShell.CurrentDirectory = App.Path
    WshShell.Run "explorer.exe /e, " & staticIntelPath
    Set WshShell = Nothing
End Sub

Private Sub btnUpdateEntry_Click()
selectedText = lstFeed.List(lstFeed.ListIndex)
pipeArray = txtURL.Text & "|" & "\cache\intel\" & LCase(txtCacheLocation.Text) & "|" & txtFeedString.Text & "|" & intToBoolStr(chkEnabled.Value) & "|" & txtRefresh.Text & "|" & intToBoolStr(ChkIgnoreSSL.Value) & "|" & txtAgeLimit & "|" & ComboCategory

    
If lstFeed.ListIndex = -1 And btnUpdateEntry.Caption = "Add New Feed With Current Values" Then
    If Len(txtURL.Text) < 11 Or Len(txtFeedString.Text) < 1 Or _
    Len(txtFeedString.Text) < 1 Then

        If Len(txtCacheLocation.Text) < 1 Then flashBackColor FrmFeed.txtCacheLocation
        If Len(txtFeedString.Text) < 1 Then flashBackColor FrmFeed.txtFeedString
        If Len(txtURL.Text) < 11 Then flashBackColor FrmFeed.txtURL
        Exit Sub
    End If
    If IsNumeric(txtIntelColumn) And IsNumeric(txtDescColumn) Then
        Form1.dictCSVFeed.Item(feedName) = txtIntelColumn.Text & "|" & txtDescColumn.Text
    ElseIf chkCSV.Value = 1 Then
        If Len(txtIntelColumn.Text) < 1 Then flashBackColor FrmFeed.txtIntelColumn
        If Len(txtDescColumn.Text) < 1 Then flashBackColor FrmFeed.txtDescColumn
        Exit Sub
    End If
    ReadFeedlist "\cache\intel\" & txtCacheLocation.Text & "|" & txtAgeLimit, pipeArray, False, True, False 'populating both args will save to dict
    
ElseIf lstFeed.ListIndex <> -1 Then
    
    If chkCSV.Value = 1 Then
        feedName = "\cache\intel\" & txtCacheLocation.Text
        feedName = Mid(feedName, InStrRev(feedName, "\") + 1, InStrRev(feedName, ".") - 1 - InStrRev(feedName, "\"))
        If IsNumeric(txtIntelColumn) And IsNumeric(txtDescColumn) Then
            Form1.dictCSVFeed.Item(feedName) = txtIntelColumn.Text & "|" & txtDescColumn.Text
        Else
            If Len(txtIntelColumn.Text) < 1 Then flashBackColor FrmFeed.txtIntelColumn
            If Len(txtDescColumn.Text) < 1 Then flashBackColor FrmFeed.txtDescColumn
            Exit Sub
        End If
    End If
    
    
    ReadFeedlist selectedText, pipeArray, False, True, False 'populating both args will save to dict
    boolSaveFeedChange = False
End If
    btnAdd.Caption = "Add New Feed"
    lstFeed.Enabled = True
    btnUpdateEntry.Caption = "Update The Selected Feed Entry With Changes"
End Sub

Sub TxtValidationFailed(objFormControl)
objFormControl.SetFocus
myBackColor = objFormControl.BackColor
objFormControl.BackColor = vbHighlight
End Sub

Private Sub chk_Attacker_Click()
Form1.boolSaveConfigPrompt = True
Form1.intelAttacker = chk_Attacker.Value
End Sub

Private Sub chk_Malware_Click()
Form1.boolSaveConfigPrompt = True
Form1.intelMalware = chk_Malware.Value
End Sub

Private Sub chk_Multi_Click()
Form1.boolSaveConfigPrompt = True
Form1.intelMulti = chk_Multi.Value
End Sub

Private Sub chk_proxy_Click()
Form1.boolSaveConfigPrompt = True
Form1.intelProxy = chk_proxy.Value
End Sub

Private Sub chk_Static_Click()
Form1.boolSaveConfigPrompt = True
Form1.staticIntel = chk_Static.Value
End Sub

Private Sub chkCSV_Click()
        If chkCSV.Value = 1 Then
            txtIntelColumn.Enabled = True
            txtDescColumn.Enabled = True
            txtIntelColumn.Visible = True
            txtDescColumn.Visible = True
            lblcolumn.Visible = True
            lblDesc.Visible = True
        Else
            txtIntelColumn.Enabled = False
            txtDescColumn.Enabled = False
            txtIntelColumn.Visible = False
            txtDescColumn.Visible = False
            lblcolumn.Visible = False
            lblDesc.Visible = False
        End If
End Sub

Private Sub chkEnabled_Click()
boolSaveFeedChange = True
End Sub

Private Sub chkWatchIntelURL_Click()
Form1.boolSaveConfigPrompt = True
Form1.WatchIntelURL = chkWatchIntelURL.Value
End Sub



Private Sub Command1_Click()

End Sub

Private Sub ComboAgeFilter_Change()
Form1.boolSaveConfigPrompt = True
Form1.ageFilter = ComboAgeFilter.Text
End Sub

Private Sub ComboAgeFilter_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub ComboCategory_Change()
populate_Cache
End Sub

Private Sub Form_Deactivate()
'MsgBox ("deactivated") 'no need to do this, why did I want to know the gui was deactivated?
End Sub

Private Sub Form_Load()
boolSaveFeedChange = False
chk_Attacker.Value = strBoolToInt(Form1.intelAttacker)

chk_Malware.Value = strBoolToInt(Form1.intelMalware)

chk_Multi.Value = strBoolToInt(Form1.intelMulti)

ComboAgeFilter.Text = Form1.ageFilter

chk_proxy.Value = strBoolToInt(Form1.intelProxy)

chk_Static.Value = strBoolToInt(Form1.staticIntel)

chkWatchIntelURL.Value = strBoolToInt(Form1.WatchIntelURL)

ComboCategory.AddItem "proxy"
ComboCategory.AddItem "malware"
ComboCategory.AddItem "attacker"
ComboCategory.AddItem "multi"
ComboCategory.ListIndex = 1
ComboAgeFilter.AddItem 1
ComboAgeFilter.AddItem 7
ComboAgeFilter.AddItem 30
ComboAgeFilter.ListIndex = 2

ReadFeedlist "", "", False, True, False 'use "" parameter to load dict

End Sub

Function ReadFeedlist(strFeedItem, pipeArray, boolSaveFeed, boolUpdateGUI, boolDelete)
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
Static dictFeeds As Scripting.Dictionary
If strFeedItem = "" And boolSaveFeed = False Then
    
    If objFSO.FileExists(App.Path & "\config\feedlist.dat") Then
    Set dictFeeds = New Scripting.Dictionary
    dictFeeds.RemoveAll
    Set objFile = objFSO.OpenTextFile(App.Path & "\config\feedlist.dat")
    Do While Not objFile.AtEndOfStream
      If Not objFile.AtEndOfStream Then 'read file
          On Error Resume Next
          FeedInfo = objFile.ReadLine
          If InStr(FeedInfo, "\cache\intel\") = 0 Then
            If InStr(FeedInfo, "cache\intel\") <> 0 Then FeedInfo = Replace(FeedInfo, "cache\intel\", "\cache\intel\")
          End If
        feedEntryArray = Split(FeedInfo, "|")
        FeedDisplay = feedEntryArray(1)
        If feedEntryArray(6) <> "" And IsNumeric(feedEntryArray(6)) Then
          
          FeedDisplay = FeedDisplay & "|" & feedEntryArray(6)

        End If
        If boolUpdateGUI = True Then lstFeed.AddItem Replace(FeedDisplay, "\cache\intel\", "")
        dictFeeds.Add FeedDisplay, FeedInfo 'sort through all aged intel after reading in entire file
      End If
    Loop
  End If
ElseIf boolDelete = True Then
    
    'need to remove from dictCSVFeed. Where is the other dict removal?
    If Form1.dictCSVFeed.Exists(lstFeed.List(lstFeed.ListIndex)) Then Form1.dictCSVFeed.Remove (lstFeed.List(lstFeed.ListIndex)) 'remove key from dict
    dictFeeds.Remove ("\cache\intel\" & lstFeed.List(lstFeed.ListIndex)) 'remove key from dict
    'remove feed from cache
    If objFSO.FileExists(App.Path & "\cache\intel\" & lstFeed.List(lstFeed.ListIndex)) Then objFSO.DeleteFile App.Path & "\config\" & lstFeed.List(lstFeed.ListIndex)
    
    lstFeed.RemoveItem (lstFeed.ListIndex) 'remove from list display
ElseIf pipeArray = "" Then 'load values
    If dictFeeds.Exists("\cache\intel\" & strFeedItem) Then
        feedEntryItems = dictFeeds.Item("\cache\intel\" & strFeedItem)
        feedEntryArray = Split(feedEntryItems, "|")
        txtURL.Text = feedEntryArray(0)
        txtCacheLocation.Text = Replace(feedEntryArray(1), "\cache\intel\", "")
        txtFeedString.Text = feedEntryArray(2)
        boolUseFeed = feedEntryArray(3)
        chkEnabled.Value = strBoolToInt(boolUseFeed)
        txtRefresh.Text = feedEntryArray(4)
        boolIgnoreSSL = feedEntryArray(5)
        ChkIgnoreSSL.Value = strBoolToInt(boolIgnoreSSL)
        txtAgeLimit = feedEntryArray(6)
        ComboCategory = feedEntryArray(7)
        
        If boolUpdateGUI = True Then
            feedName = cache2FeedName(txtCacheLocation.Text)
            
         
            'clear gui elements for feed csv intel columns
            chkCSV.Value = 0
            txtIntelColumn.Text = ""
            txtDescColumn.Text = ""
            If Form1.dictCSVFeed.Exists(feedName) Then 'load CSV intel columns
            reporValues = Form1.dictCSVFeed.Item(feedName)
                If InStr(reporValues, "|") > 0 Then
                      arrayRval = Split(reporValues, "|")
                  txtIntelColumn.Text = arrayRval(0)
                  txtDescColumn.Text = arrayRval(1)
                  chkCSV.Value = 1
                End If
            End If
        End If
        
    End If
ElseIf pipeArray <> "" Then 'store values
    feedEntryArray = Split(pipeArray, "|")
    FeedDisplay = feedEntryArray(1)
    feedName = cache2FeedName(FeedDisplay)
    Form1.dictCSVFeed.Item(feedName) = txtIntelColumn.Text & "|" & txtDescColumn.Text
    
    
    If feedEntryArray(6) <> "" And IsNumeric(feedEntryArray(6)) Then
          FeedDisplay = FeedDisplay & "|" & feedEntryArray(6)
    End If
    dictFeeds.Item(FeedDisplay) = pipeArray 'save entry in dict
    If lstFeed.ListIndex = -1 Then
        lstFeed.AddItem (Replace(FeedDisplay, "\cache\intel\", ""))
        lstFeed.ListIndex = lstFeed.ListCount - 1
    Else 'Update with new values
        If FeedDisplay <> "\cache\intel\" & lstFeed.List(lstFeed.ListIndex) Then
            If Form1.dictCSVFeed.Exists(lstFeed.List(lstFeed.ListIndex)) Then Form1.dictCSVFeed.Remove (lstFeed.List(lstFeed.ListIndex))  'remove key from dict
            
            If objFSO.FileExists(App.Path & "\cache\intel\" & lstFeed.List(lstFeed.ListIndex)) Then objFSO.DeleteFile App.Path & "\config\" & lstFeed.List(lstFeed.ListIndex)

            dictFeeds.Remove ("\cache\intel\" & lstFeed.List(lstFeed.ListIndex)) 'remove old key from dict
            lstFeed.List(lstFeed.ListIndex) = Replace(FeedDisplay, "\cache\intel\", "") 'replace list entry with new key
        End If
    End If
End If

If boolSaveFeed = True Then

'save feeds
If objFSO.FileExists(App.Path & "\config\feedlist.dat") Then objFSO.DeleteFile App.Path & "\config\feedlist.dat"
Set f = objFSO.OpenTextFile(App.Path & "\config\feedlist.dat", ForAppending, True)
For Each feedEntry In dictFeeds 'dict feeds need to stay
f.WriteLine dictFeeds.Item(feedEntry) 'URL|txtCacheLocation|txtFeedString|boolUseFeed|txtRefresh.Text|boolIgnoreSSL|txtAgeLimit|ComboCategory
Next
f.Close
End If


End Function

Function cache2FeedName(cachePath)

If InStrRev(cachePath, "\") > 0 And InStrRev(cachePath, ".") > 0 Then
            cache2FeedName = Mid(cachePath, InStrRev(cachePath, "\") + 1, InStrRev(cachePath, ".") - 1 - InStrRev(cachePath, "\"))
ElseIf InStrRev(cachePath, ".") > 0 Then
    cache2FeedName = Left(cachePath, InStrRev(cachePath, ".") - 1)
Else
    cache2FeedName = cachePath
End If
End Function


Function strBoolToInt(strBool)
If LCase(strBool) = "false" Then
    strBoolToInt = 0
ElseIf LCase(strBool) = "true" Then
    strBoolToInt = 1
Else
MsgBox "Failed to convert string to int:" & strBool
End If

End Function

Function intToBoolStr(intBool)
If intBool = 0 Then
    intToBoolStr = "False"
ElseIf intBool = 1 Then
    intToBoolStr = "True"
Else
MsgBox "Failed to convert int to boolean string:" & CStr(intBool)
End If

End Function

Private Sub Form_Terminate()
'application termination not form close
End Sub

Private Sub Form_Unload(Cancel As Integer)
If boolSaveFeedChange = True Then
    feedAnswer = MsgBox("Would you like to save changes to the current feed?", vbYesNo, "VTTL - Feed Change")
    If feedAnswer = vbYes Then
        btnUpdateEntry_Click
    Else
        boolSaveFeedChange = False
    End If
End If
End Sub

Private Sub lstFeed_Click()
'Want to add this, but need to check last index and save that not this one.
'If boolSaveFeedChange = True Then
'    feedAnswer = MsgBox("Would you like to save changes to the last selected feed?", vbQuestion, "VTTL - Feed Change")
'    If feedAnswer = vbYes Then
'        btnUpdateEntry_Click
'    Else
'        boolSaveFeedChange = False
'    End If
'End If

If lstFeed.ListIndex = -1 Then Exit Sub
selectedText = lstFeed.List(lstFeed.ListIndex)
ReadFeedlist selectedText, "", False, True, False
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub lstFeed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    tmpAnswer = MsgBox("Are you sure you want to delete the feed?", vbYesNo, "Question")
    If tmpAnswer = vbYes Then
        selectedText = lstFeed.List(lstFeed.ListIndex)
        ReadFeedlist selectedText, "", False, True, True
    End If
End If
End Sub

Private Sub lstFeed_KeyPress(KeyAscii As Integer)
'MsgBox (KeyAscii)
End Sub

Private Sub timerFlashy_Timer()
Static AlternatingCount As Integer
If AlternatingCount = 6 Then
    timerFlashy.Enabled = False
    AlternatingCount = 0
    flashBackColor New FileSystemObject
    Exit Sub
End If
AlternatingCount = 1 + AlternatingCount
flashBackColor Nothing
End Sub

Function flashBackColor(objFormControl)
Static CurrentObject As Object
Static currentBackColor As Variant
If CurrentObject Is Nothing And objFormControl Is Nothing Then Exit Function
If CurrentObject Is Nothing Then
Set CurrentObject = objFormControl
    currentBackColor = objFormControl.BackColor
    CurrentObject.SetFocus
    timerFlashy.Enabled = True
ElseIf objFormControl Is Nothing Then
    If CurrentObject.BackColor = vbHighlight Then
        CurrentObject.BackColor = currentBackColor
    Else
        CurrentObject.BackColor = vbHighlight
    End If
ElseIf TypeName(objFormControl) <> TypeName(New FileSystemObject) Then
    Set CurrentObject = objFormControl
    currentBackColor = objFormControl.BackColor
    CurrentObject.SetFocus
    timerFlashy.Enabled = True
End If


End Function

Private Sub txtAgeLimit_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtDescColumn_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtFeedString_LostFocus()
populate_Cache
End Sub

Private Sub txtIntelColumn_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub

Sub populate_Cache()
If txtURL.Text <> "" And txtCacheLocation.Text = "" Then
    If Len(txtCacheLocation.Text) = 0 Then
    tmpFileName = Right(txtURL.Text, Len(txtURL.Text) - InStrRev(txtURL.Text, "/"))
    txtCacheLocation.Text = tmpFileName
    End If
End If
End Sub

Private Sub txtURL_LostFocus()
populate_Cache
If Len(txtURL.Text) > 4 Then
    If LCase(Right(txtURL.Text, 4)) = ".csv" Then
        chkCSV.Value = 1
    End If
End If
End Sub
