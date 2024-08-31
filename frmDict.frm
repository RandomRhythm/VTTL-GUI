VERSION 5.00
Begin VB.Form frmDict
   Caption         =   "Dictionary Editor"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13830
   Icon            =   "frmDict.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   11055
   ScaleWidth      =   13830
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   22
      Left            =   9120
      TabIndex        =   50
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   21
      Left            =   6960
      TabIndex        =   49
      Top             =   8640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   20
      Left            =   10920
      TabIndex        =   48
      Top             =   8160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Avira"
      Height          =   495
      Index           =   11
      Left            =   3960
      TabIndex        =   47
      Top             =   8640
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Hacker Tool"
      Height          =   495
      Index           =   23
      Left            =   10920
      TabIndex        =   42
      Top             =   8640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   24
      Left            =   840
      TabIndex        =   31
      Top             =   9600
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   25
      Left            =   2640
      TabIndex        =   32
      Top             =   9600
      Width           =   2055
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   26
      Left            =   840
      TabIndex        =   33
      Top             =   10080
      Width           =   1695
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   27
      Left            =   2640
      TabIndex        =   34
      Top             =   10080
      Width           =   1935
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   5
      Left            =   3960
      TabIndex        =   15
      Top             =   7680
      Width           =   1455
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   7
      Left            =   2640
      TabIndex        =   17
      Top             =   8160
      Width           =   1335
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   4
      Left            =   2640
      TabIndex        =   14
      Top             =   7680
      Width           =   1335
   End
   Begin VB.OptionButton optDict 
      Caption         =   "ESET"
      Height          =   495
      Index           =   1
      Left            =   2640
      TabIndex        =   11
      Top             =   7200
      Width           =   1335
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Bitdefender"
      Height          =   495
      Index           =   9
      Left            =   840
      TabIndex        =   19
      Top             =   8640
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Avira"
      Height          =   495
      Index           =   10
      Left            =   2640
      TabIndex        =   20
      Top             =   8640
      Width           =   2175
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Dr.Web"
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   10
      Top             =   7200
      Width           =   1215
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   7200
      Width           =   1695
   End
   Begin VB.OptionButton optDict 
      Caption         =   "klj"
      Height          =   495
      Index           =   3
      Left            =   840
      TabIndex        =   13
      Top             =   7680
      Width           =   1335
   End
   Begin VB.OptionButton optDict 
      Caption         =   "lkj"
      Height          =   495
      Index           =   6
      Left            =   840
      TabIndex        =   16
      Top             =   8160
      Width           =   1335
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   8
      Left            =   3960
      TabIndex        =   18
      Top             =   8160
      Width           =   1575
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   15
      Left            =   6960
      TabIndex        =   24
      Top             =   7680
      Width           =   2175
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   19
      Left            =   9120
      TabIndex        =   28
      Top             =   8160
      Width           =   1935
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   13
      Left            =   9120
      TabIndex        =   22
      Top             =   7200
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   18
      Left            =   6960
      TabIndex        =   27
      Top             =   8160
      Width           =   2055
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Microsoft"
      Height          =   495
      Index           =   12
      Left            =   6960
      TabIndex        =   21
      Top             =   7200
      Width           =   2055
   End
   Begin VB.OptionButton optDict 
      Caption         =   "McAfee"
      Height          =   495
      Index           =   16
      Left            =   9120
      TabIndex        =   25
      Top             =   7680
      Width           =   1695
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   14
      Left            =   10920
      TabIndex        =   23
      Top             =   7200
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   495
      Index           =   17
      Left            =   10920
      TabIndex        =   26
      Top             =   7680
      Width           =   1815
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   375
      Index           =   29
      Left            =   9120
      TabIndex        =   30
      Top             =   9600
      Width           =   2295
   End
   Begin VB.OptionButton optDict 
      Caption         =   "Option1"
      Height          =   375
      Index           =   28
      Left            =   6960
      TabIndex        =   29
      Top             =   9600
      Width           =   2295
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Search"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "Watchlists"
      Height          =   1335
      Left            =   720
      TabIndex        =   43
      Top             =   9360
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Exclusions"
      Height          =   735
      Left            =   6840
      TabIndex        =   41
      Top             =   9360
      Width           =   4935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Intelligence"
      Height          =   2295
      Left            =   6840
      TabIndex        =   40
      Top             =   6960
      Width           =   6255
   End
   Begin VB.Frame Frame1 
      Caption         =   "Vendors"
      Height          =   2295
      Left            =   720
      TabIndex        =   39
      Top             =   6960
      Width           =   5175
   End
   Begin VB.CommandButton bntPathPrev 
      Caption         =   "Edit Path Vendor Combiation"
      Height          =   495
      Left            =   10080
      TabIndex        =   37
      ToolTipText     =   "Prevalence of file path and reported vendor name combination"
      Top             =   10200
      Width           =   2895
   End
   Begin VTTL_GUI.ListView listDisplay 
      Height          =   5175
      Left            =   360
      TabIndex        =   9
      Top             =   1680
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   9128
      View            =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HideSelection   =   0   'False
      ShowInfoTips    =   -1  'True
      ShowLabelTips   =   -1  'True
      PictureAlignment=   5
   End
   Begin VB.TextBox txtInt 
      Height          =   495
      Left            =   12480
      TabIndex        =   38
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton btnUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton btnGryware 
      Caption         =   "Digital Signature Grayware List"
      Height          =   495
      Left            =   6840
      TabIndex        =   36
      Top             =   10200
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset File"
      Height          =   495
      Left            =   7920
      TabIndex        =   35
      Top             =   1080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "Edit File"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5280
      TabIndex        =   7
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   495
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Width           =   7575
   End
   Begin VB.TextBox txtKey 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.ListBox ListValues 
      Height          =   5130
      Left            =   4680
      TabIndex        =   0
      Top             =   1680
      Width           =   7575
   End
   Begin VB.Label Label_3 
      Caption         =   "Label1"
      Height          =   255
      Left            =   12480
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label_1 
      Height          =   255
      Left            =   360
      TabIndex        =   45
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label_2 
      Height          =   255
      Left            =   4680
      TabIndex        =   44
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmDict"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194
Dim intIndexEdit
Dim intOptionSelectedIndex
Dim boolRestrictToInteger
Const ForReading = 1

Function ReturnFilePath(strOption)
Dim returnFpath

txtValue.Visible = True
txtInt.Visible = False
Select Case strOption
        Case "AlienVault NIDS Category"
            returnFpath = "\cache\NIDS_Cat.dat"
        Case "AlienVault Signatures"
            returnFpath = "\cache\NIDS_Sig.dat"
        Case "Avira"
            returnFpath = "\cache\avira.dat"
        Case "Dr.Web"
            returnFpath = "\cache\drweb.dat"
        Case "Digital Signature"
            returnFpath = "\cache\digsig.dat"

        Case "Digital Signer Associated Websites"
            returnFpath = "\cache\orgwho.dat"
        Case "ESET"
            returnFpath = "\cache\eset.dat"

        Case "Hacker Tools"
            returnFpath = "\cache\hktl.dat"
            boolRestrictToInteger = True
        Case "IP/Domain"
            returnFpath = "\config\IPDwatchlist.txt"
        Case "Malware Hash"
            returnFpath = "\config\malhash.dat"
        Case "McAfee"
            returnFpath = "\cache\mcafee.dat"
        Case "Microsoft"
            returnFpath = "\cache\microsoft.dat"
        Case "Panda"
            returnFpath = "\cache\panda.dat"
        Case "Family Names"
            returnFpath = "\cache\family.dat"
            listDisplay.Visible = True
            'txtInt.Visible = True
        Case "Potentially Unwanted"
            returnFpath = "\cache\pup.dat"
            boolRestrictToInteger = True

        Case "Potentially Unwanted Digital Signature"
            returnFpath = "\cache\ds_pup.dat"
            boolRestrictToInteger = True
        Case "Sophos"
            returnFpath = "\cache\sophos.dat"
        Case "Symantec"
            returnFpath = "\cache\symantec.dat"
        Case "TrendMicro"
            returnFpath = "\cache\trendmicro.dat"
        Case "Hash Whitelist"
            returnFpath = "\config\whitehash.dat"
        Case "Detection Name"
            returnFpath = "\config\DNwatchlist.txt"
            txtValue.Visible = False
        Case "IP/Domain"
            returnFpath = "\config\IPDwatchlist.txt"
        Case "Keyword"
            returnFpath = "\config\KWordwatchlist.txt"
            txtValue.Visible = False
        Case "URL"
            returnFpath = "\config\URLwatchlist.txt"
            txtValue.Visible = False
        Case "Subdomain No Submit"
            txtValue.Visible = False
            writeNoSubmitSubdomains
            returnFpath = "\config\VTTL_domains.txt"
        Case "Domain/IP No Submit"
            txtValue.Visible = False
            writeNoSubmitExact
            returnFpath = "\config\VTTL_NoSubmit.txt"
    End Select
ReturnFilePath = App.Path & returnFpath
End Function



Private Sub bntPathPrev_Click()
Shell "notepad" & " " & Chr(34) & App.Path & "\cache\pathvend.dat" & Chr(34), vbNormalFocus
End Sub

Private Sub btnAdd_Click()
'listDisplay.AddItem txtKey.Text
'listDisplay.AddItem txtValue.Text
Set li = listDisplay.ListItems.Add(, , txtKey.Text)

If listDisplay.ColumnHeaders.Count = 0 Then
     If txtKey.Visible = True Then listDisplay.ColumnHeaders.Add , , "Item"  ', listDisplay.Width / 2
     If txtValue.Visible = True Then listDisplay.ColumnHeaders.Add , , "Value"
End If

If txtValue.Visible = True Then
    li.SubItems(1) = txtValue.Text
End If
End Sub

Private Sub btnDelete_Click()
If listDisplay.SelectedItem.Index <> -1 Then
    listDisplay.ListItems.Remove (listDisplay.SelectedItem.Index)
    'listDisplay.RemoveItem (intIndexEdit)
End If
End Sub

Private Sub btnEdit_Click()
strDictPath = ReturnFilePath(optDict(intOptionSelectedIndex).Caption)
Shell "notepad" & " " & Chr(34) & strDictPath & Chr(34), vbNormalFocus
End Sub

Private Sub btnGryware_Click()
        'Case "Digital Sig Potentially Unwanted"
         '   returnFpath = "\cache\ds_gry.dat"
         Shell "notepad" & " " & Chr(34) & App.Path & "\cache\ds_gry.dat" & Chr(34), vbNormalFocus
End Sub

Public Sub SetListboxScrollbar(ByVal lst As ListBox)
Dim i As Integer
Dim new_len As Long
Dim max_len As Long

    For i = 0 To lst.ListCount - 1
        new_len = 10 + lst.Parent.ScaleX( _
            lst.Parent.TextWidth(lst.List(i)), _
            lst.Parent.ScaleMode, vbPixels)
        If max_len < new_len Then max_len = new_len
    Next i

    SendMessage lst.hWnd, _
        LB_SETHORIZONTALEXTENT, _
        max_len, 0
End Sub
 


Private Sub btnSave_Click()
Set savefso = CreateObject("Scripting.FileSystemObject")
Dim SaveFilePath
SaveFilePath = ReturnFilePath(optDict(intOptionSelectedIndex).Caption)
If savefso.FileExists(SaveFilePath) Then
    If savefso.FileExists(SaveFilePath & ".bak") Then
        savefso.DeleteFile SaveFilePath & ".bak"
    End If
    If savefso.FileExists(SaveFilePath & ".bak") Then
        MsgBox "Unable to remove backup file: " & SaveFilePath & ".bak. Will not proceed with saving."
        Exit Sub
    Else
        savefso.MoveFile SaveFilePath, SaveFilePath & ".bak"
    End If
End If
If savefso.FileExists(SaveFilePath) Then
    MsgBox "Unable to backup file: " & SaveFilePath & ". Will not proceed with saving."
    Exit Sub
End If
If txtValue.Visible = True Then
    For Each ListItem In listDisplay.ListItems
        Form1.LogData CStr(SaveFilePath), CStr(ListItem) & "|" & CStr(ListItem.SubItems(1))
    
    Next
Else
    For Each ListItem In listDisplay.ListItems
        Form1.LogData CStr(SaveFilePath), CStr(ListItem)
    Next
End If
End Sub

Private Sub btnSearch_Click()
Searchlist frmDict.txtKey.Text, False
End Sub

Sub Searchlist(strSearchText, boolPartial)

Dim intTmpIndex As Long
Dim mytestitem As LvwListItem
intTmpIndex = 0
Set mytestitem = listDisplay.FindItem(strSearchText, intTmpIndex, boolPartial, True)

If Not mytestitem Is Nothing Then
    mytestitem.Selected = True
    ListDiplayLoadItems
ElseIf boolPartial = False Then
    Searchlist strSearchText, True
End If

End Sub


Private Sub btnUpdate_Click()
If listDisplay.SelectedItem Is Nothing Then
Exit Sub
End If
If listDisplay.SelectedItem.Index <> -1 Then
    listDisplay.ListItems(listDisplay.SelectedItem.Index).Text = txtKey.Text
    If txtValue.Visible = True Then
        listDisplay.ListItems(listDisplay.SelectedItem.Index).SubItems(1) = txtValue.Text
    End If
    If txtInt.Visible = True Then

        listDisplay.ListItems(listDisplay.SelectedItem.Index).SubItems(2) = txtInt.Text

    End If
End If
End Sub



Private Sub Form_Load()
frmDict.Caption = "Dictionary Editor - " & App.Path
For IntOptionCount = 0 To optDict.Count
    Select Case IntOptionCount
        Case 0
            optDict(IntOptionCount).Caption = "Avira"
        Case 1
            optDict(IntOptionCount).Caption = "Dr.Web"
        Case 2
            optDict(IntOptionCount).Caption = "ESET"
            optDict(IntOptionCount).ToolTipText = "Hacker tool names and score to help with categorization"
        Case 3
            optDict(IntOptionCount).Caption = "Microsoft"
        Case 4
            optDict(IntOptionCount).Caption = "McAfee"
        Case 5
            optDict(IntOptionCount).Caption = "Panda"
        Case 6
            optDict(IntOptionCount).Caption = "Sophos"
        Case 7
            optDict(IntOptionCount).Caption = "Symantec"
        Case 8
            optDict(IntOptionCount).Caption = "TrendMicro"
            optDict(IntOptionCount).ToolTipText = "Add to adjusted malware score for custom list malware and update detection name if one was given in malhash.dat"
        Case 9
            optDict(IntOptionCount).Caption = "AlienVault Signatures"
            optDict(IntOptionCount).ToolTipText = "Requires payed subscription"
        Case 10
            optDict(IntOptionCount).Caption = "AlienVault NIDS Category"
            optDict(IntOptionCount).ToolTipText = "Requires payed subscription"
        Case 11

        Case 12
            optDict(IntOptionCount).Caption = "Hash Whitelist"
            optDict(IntOptionCount).ToolTipText = "Adjusted malware score to zero and update detection name if one was given in whitehash.dat"
        Case 13
            optDict(IntOptionCount).Caption = "Malware hash"
            optDict(IntOptionCount).ToolTipText = "Increase malicious score and report associated detection name"


        Case 14
            optDict(IntOptionCount).Caption = "Family Names"
            optDict(IntOptionCount).ToolTipText = "Common names used to identify family name"
            
        Case 15
            optDict(IntOptionCount).Caption = "Digital Signer Associated Websites"
            optDict(IntOptionCount).ToolTipText = "Domains registered with the same name as the digital signer"
        Case 16
            optDict(IntOptionCount).Caption = "Hacker Tools"
            optDict(IntOptionCount).ToolTipText = "Hacker tool names and score to help with categoriztion"
        Case 17
                    optDict(IntOptionCount).Caption = "Potentially Unwanted"
            optDict(IntOptionCount).ToolTipText = "List of potentially unwanted applications"
        Case 18
            optDict(IntOptionCount).Caption = "Digital Signature"
            optDict(IntOptionCount).ToolTipText = "Digital signer prevalence"
            
        Case 19
            
            optDict(IntOptionCount).Caption = "Potentially Unwanted Digital Signature"
            optDict(IntOptionCount).ToolTipText = "Signers of potentially unwanted applications"
            
            
        Case 24
            optDict(IntOptionCount).Caption = "Detection Name"
            optDict(IntOptionCount).ToolTipText = "Detection name labels to watch for provided in lowercase"
        Case 25
            optDict(IntOptionCount).Caption = "IP/Domain"
            optDict(IntOptionCount).ToolTipText = "Dictionary of domains and IP addresses and their association"
        Case 26
            optDict(IntOptionCount).Caption = "Keyword"
            optDict(IntOptionCount).ToolTipText = "Keywords to watch for in AlienVault OTX pulses"
        Case 27
            optDict(IntOptionCount).Caption = "URL"
            optDict(IntOptionCount).ToolTipText = "URL structure watchlist. Regex support is configured in INI (UseRegexForURL)"
            
        Case 28
            optDict(IntOptionCount).Caption = "Subdomain No Submit"
            optDict(IntOptionCount).ToolTipText = "Do not submit sub domains to vendors. Format is .domaintoexclude.com"
        Case 29
            optDict(IntOptionCount).Caption = "Domain/IP No Submit"
            optDict(IntOptionCount).ToolTipText = "Do not submit these domains or IP addresses (exact match) to vendors"
    End Select
Next

End Sub

Sub writeNoSubmitExact()
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(App.Path & "\VTTL_NoSubmit.txt") = False Then
    Form1.LogData App.Path & "\VTTL_NoSubmit.txt", "ocsp.int-x3.letsencrypt.org"
End If
End Sub
Sub writeNoSubmitSubdomains()
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(App.Path & "\VTTL_domains.txt") = False Then
  Form1.LogData App.Path & "\VTTL_domains.txt", ".local" 'exclude .local DNS names from lookups
  'exclude other possible corporate utilized domain names
  Form1.LogData App.Path & "\VTTL_domains.txt", ".my.carbonblack.io"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".onmicrosoft.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".okta.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".sharepoint.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".service-now.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".silkroad.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".salesforce.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".force.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".allegiancetech.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".account.box.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".corporateperks.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".bamboohr.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".acquia-sites.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".fbcdn.net"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".cbssports.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".atlassian.net"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".zoom.us"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".qualtrics.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".dynamics.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".ultipro.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".sophosxl.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".footprintdns.com"
  Form1.LogData App.Path & "\VTTL_domains.txt", ".barracudabrts.com"

End If
End Sub



Private Sub listDisplay_Click()
ListDiplayLoadItems
End Sub

Sub ListDiplayLoadItems()
Dim tmpListArray

If listDisplay.ListItems.Count > 0 Then
    txtKey.Text = listDisplay.SelectedItem.Text
    If txtValue.Visible = True Then
        txtValue.Text = listDisplay.SelectedItem.SubItems(1)
    End If
    If txtInt.Visible = True Then
        txtInt.Text = listDisplay.SelectedItem.SubItems(2)
    End If
End If
End Sub

Private Sub optDict_Click(Index As Integer)
intOptionSelectedIndex = Index
intIndexEdit = -1
strDictPath = ReturnFilePath(optDict(Index).Caption)
txtKey.Text = ""
txtValue.Text = ""
If txtValue.Visible = True Then
    loadDict strDictPath
Else
    loadList strDictPath
End If

If Index < 9 Then
    Label_1.Caption = "Detection Name"
    Label_2.Caption = "Reference"
ElseIf Index = 9 Then
    Label_1.Caption = "ID"
    Label_2.Caption = "Signature"
ElseIf Index = 10 Then
    Label_1.Caption = "ID"
    Label_2.Caption = "Category"

ElseIf Index = 12 Then
    Label_1.Caption = "Hash"
    Label_2.Caption = "Whitelist text"
ElseIf Index = 13 Then
    Label_1.Caption = "Hash"
    Label_2.Caption = "Malicious association text"
ElseIf Index = 14 Then
    Label_1.Caption = "Family name"
    Label_2.Caption = "Change to family name"
ElseIf Index = 15 Then
    Label_1.Caption = "Signer"
    Label_2.Caption = "Websites with ^ separated"
ElseIf Index = 16 Then
    Label_1.Caption = "Hack tool name"
    Label_2.Caption = "Hack tool score"
ElseIf Index = 17 Then
    Label_1.Caption = "Potentially unwanted application name"
    Label_2.Caption = "PUA score"
ElseIf Index = 18 Then
    Label_1.Caption = "Digital signature"
    Label_2.Caption = "Prevalence"
ElseIf Index = 19 Then
    Label_1.Caption = "Digital signature"
    Label_2.Caption = "PUA score"
End If


If optDict(Index).Caption = "Detection Name" Then
    Label_1.Caption = "Detection Name"
    Label_2.Caption = ""
ElseIf optDict(Index).Caption = "IP/Domain" Then
    Label_1.Caption = "IP/Domain Name"
    Label_2.Caption = "Comment"
ElseIf optDict(Index).Caption = "Keyword" Then
    Label_1.Caption = "Keyword"
    Label_2.Caption = ""
ElseIf optDict(Index).Caption = "URL" Then
    Label_1.Caption = "URL"
    Label_2.Caption = ""
ElseIf optDict(Index).Caption = "Subdomain No Submit" Then
    Label_1.Caption = "IP/Domain Name"
    Label_2.Caption = ""
ElseIf optDict(Index).Caption = "Domain/IP No Submit" Then
    Label_1.Caption = "IP/Domain Name"
    Label_2.Caption = ""
End If

End Sub
Sub loadList(strDictionaryPath)
Set objFSO = CreateObject("Scripting.FileSystemObject")
listDisplay.ListItems.Clear
listDisplay.ColumnHeaders.Clear
If objFSO.FileExists(strDictionaryPath) = False Then
    Exit Sub
End If
Set objTextFile = objFSO.OpenTextFile _
(strDictionaryPath, ForReading)
If objTextFile.AtEndOfStream = False Then
    Do While objTextFile.AtEndOfStream = False
        
        strTmpLine = objTextFile.ReadLine
        
            If boolViewHeader = False Then

                    listDisplay.ColumnHeaders.Add , , "Item" ', listDisplay.Width / 2
                    listDisplay.ColumnHeaders.Add , , "Value"
                    intHeaderCount = 1

                boolViewHeader = True
            End If
                        Set li = listDisplay.ListItems.Add(, , strTmpLine)
            li.SubItems(1) = ""
            
        
    Loop
End If
If boolViewHeader = False Then 'not having a header causes a crash so we need to ensure there is one
    'loaded lists that have no content don't get a header created above so creating here
    listDisplay.ColumnHeaders.Add , , "Item" ', listDisplay.Width / 2
    listDisplay.ColumnHeaders.Add , , "Value" ', listDisplay.Width / 2
    intHeaderCount = 1
    boolViewHeader = True
End If
End Sub

Sub loadDict(strDictionaryPath)
SetupVisualStyles Me
'need to make sure this doesn't lock the file. Causes VTTL to crash when it can't get exclusive access.
Dim boolViewHeader: boolViewHeader = False
Dim intHeaderCount
btnSave.Enabled = True
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strDictionaryPath) = False Then
    If Right(strDictionaryPath, 11) = "malhash.dat" Then
        Form1.LogData CStr(strDictionaryPath), "44d88612fea8a8f36de82e1278abb02f|eicar_test_file"
    End If
    If Right(strDictionaryPath, 13) = "whitehash.dat" Then
        Form1.LogData CStr(strDictionaryPath), "44d88612fea8a8f36de82e1278abb02f|eicar_test_file"
    End If

End If


listDisplay.ListItems.Clear
listDisplay.ColumnHeaders.Clear
If objFSO.FileExists(strDictionaryPath) = False Then
    Exit Sub
End If
Set objTextFile = objFSO.OpenTextFile _
(strDictionaryPath, ForReading)
If objTextFile.AtEndOfStream = False Then
    Do While objTextFile.AtEndOfStream = False
        
        strTmpLine = objTextFile.ReadLine
        If InStr(strTmpLine, "|") > 0 Then
            arrayline = Split(strTmpLine, "|")
            If boolViewHeader = False Then
                If UBound(arrayline) > 1 Then
                    listDisplay.ColumnHeaders.Add , , "File Path" ', listDisplay.Width / 3
                    listDisplay.ColumnHeaders.Add , , "Publisher" ', listDisplay.Width / 3
                    listDisplay.ColumnHeaders.Add , , "Prevalence" ', listDisplay.Width / 3
                    intHeaderCount = 2
                Else
                    listDisplay.ColumnHeaders.Add , , "Item" ', listDisplay.Width / 2
                    listDisplay.ColumnHeaders.Add , , "Value" ', listDisplay.Width / 2
                    intHeaderCount = 1
                End If
                boolViewHeader = True
            End If
            Set li = listDisplay.ListItems.Add(, , arrayline(0))
            li.SubItems(1) = arrayline(1)

            If intHeaderCount > 1 Then
                    li.SubItems(2) = arrayline(2)
            End If
        End If
    Loop
End If
objTextFile.Close
If boolViewHeader = False Then 'not having a header causes a crash so we need to ensure there is one
    'loaded lists that have no content don't get a header created above so creating here
    listDisplay.ColumnHeaders.Add , , "Item" ', listDisplay.Width / 2
    listDisplay.ColumnHeaders.Add , , "Value" ', listDisplay.Width / 2
    intHeaderCount = 1
    boolViewHeader = True
End If
SetListboxScrollbar ListValues

End Sub

Private Sub VScroll2_Change()

End Sub


Private Sub txtKey_Change()

End Sub

Private Sub txtKey_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    btnSearch_Click
End If
End Sub
