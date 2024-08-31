VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VTTL INI Settings"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8220
   Icon            =   "VTTL_GUI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkShodan 
      Caption         =   "Shodan InternetDB"
      Height          =   495
      Left            =   360
      TabIndex        =   28
      ToolTipText     =   "No API key required"
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CheckBox chkCBC 
      Caption         =   "Carbon Black Enterprise EDR"
      Height          =   495
      Left            =   360
      TabIndex        =   19
      ToolTipText     =   "Formally Threat Hunter"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox chkDeepIOC 
      Caption         =   "Deep IOC Match"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      ToolTipText     =   "Try to match IOCs against associated items"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtFeed 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Feed configuration"
      ToolTipText     =   "Configure threat intel feeds"
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox lblWhois 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Whois & location configuration"
      ToolTipText     =   "Various ways to populate whois column with owner"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CheckBox chkPulsedive 
      Caption         =   "Pulsedive"
      Height          =   495
      Left            =   360
      TabIndex        =   27
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CheckBox chkDomainBLchecks 
      Caption         =   "Disable Domain block list checks"
      Height          =   255
      Left            =   3600
      TabIndex        =   50
      Top             =   4560
      Width           =   2895
   End
   Begin VB.CheckBox chkDomainCache 
      Caption         =   "Cache domain results"
      Height          =   195
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Cache API results for domains"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CheckBox chkCIF 
      Caption         =   "Collective Intelligence Framework (CIF)"
      Height          =   495
      Left            =   360
      TabIndex        =   23
      ToolTipText     =   "Collective Intelligence Framework"
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CheckBox chkWhoIsCache 
      Caption         =   "Cache whois results"
      Height          =   195
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   "Whois cache is separate from other cache"
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtSampleCount 
      Height          =   285
      Left            =   5880
      TabIndex        =   58
      Text            =   "0"
      ToolTipText     =   "Set to zero to disable category reporting"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtDetectThreshold 
      Height          =   285
      Left            =   5880
      TabIndex        =   59
      Text            =   "9"
      ToolTipText     =   "Number of positive detections"
      Top             =   5880
      Width           =   495
   End
   Begin VB.ComboBox ComboVendor 
      Height          =   315
      Left            =   5280
      TabIndex        =   32
      Text            =   "Vendor Name"
      Top             =   2040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox ChkURLregex 
      Caption         =   "URL regex parsing"
      Height          =   495
      Left            =   360
      TabIndex        =   29
      ToolTipText     =   "Use regex for URL watchlist"
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buttons"
      Height          =   1815
      Left            =   720
      TabIndex        =   49
      Top             =   6480
      Width           =   6615
      Begin VB.CommandButton btnFolder 
         Caption         =   "Output Folder"
         Height          =   495
         Left            =   4800
         TabIndex        =   67
         ToolTipText     =   "Folder directory where output is written"
         Top             =   1080
         Width           =   1215
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload Settings"
         Height          =   495
         Left            =   480
         TabIndex        =   65
         ToolTipText     =   "Loads settings from INI"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CommandButton cmdDictEdit 
         Caption         =   "Dictionary and List Editor"
         Height          =   495
         Left            =   2520
         TabIndex        =   66
         ToolTipText     =   "Advanced settings"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton btnScanList 
         Caption         =   "Preview Scan List"
         Height          =   495
         Left            =   240
         TabIndex        =   61
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton cmdEditScanList 
         Caption         =   "Edit Scan List"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Edit text file containing the list of hashes or domain/IP addresses to lookup"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Lookups"
         Height          =   495
         Left            =   3600
         TabIndex        =   63
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton saveBtn 
         Caption         =   "Save Config"
         Height          =   495
         Left            =   5160
         TabIndex        =   64
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkMalshare 
      Caption         =   "Malshare API"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.OptionButton optHash 
      Caption         =   "Domain/IP Address"
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   11
      Top             =   1320
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton optHash 
      Caption         =   "Hash"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtBarracuda 
      Height          =   285
      Left            =   5880
      TabIndex        =   45
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CheckBox chkBarracuda 
      Caption         =   "Barracuda"
      Height          =   255
      Left            =   3600
      TabIndex        =   44
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CheckBox chkTIA 
      Caption         =   "ThreatIntelligenceAggregator"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      ToolTipText     =   "threatintelligenceaggregator.org"
      Top             =   2280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkBLchecks 
      Caption         =   "Enable block list checks"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.TextBox txtSORBS 
      Height          =   285
      Left            =   5880
      TabIndex        =   43
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CheckBox chkSORBS 
      Caption         =   "Enable SORBS"
      Height          =   495
      Left            =   3600
      TabIndex        =   42
      Top             =   3720
      Width           =   1935
   End
   Begin VB.TextBox txtSURBL 
      Height          =   285
      Left            =   5880
      TabIndex        =   41
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CheckBox chkSURBL 
      Caption         =   "Enable SURBL"
      Height          =   495
      Left            =   3600
      TabIndex        =   40
      ToolTipText     =   "< 250,000 queries per day"
      Top             =   3360
      Width           =   1815
   End
   Begin VB.TextBox txtZenDBL 
      Height          =   285
      Left            =   5880
      TabIndex        =   39
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtCBL 
      Height          =   285
      Left            =   5880
      TabIndex        =   37
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CheckBox chkZDBL 
      Caption         =   "Enable Zen DBL"
      Height          =   495
      Left            =   3600
      TabIndex        =   38
      ToolTipText     =   "Domain block list"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.TextBox txtURIBL 
      Height          =   285
      Left            =   5880
      TabIndex        =   35
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkCBL 
      Caption         =   "Enable cbl.abuseat.org"
      Height          =   495
      Left            =   3600
      TabIndex        =   36
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CheckBox chkURIBL 
      Caption         =   "Enable URIBL"
      Height          =   495
      Left            =   3600
      TabIndex        =   34
      ToolTipText     =   "Domain block list"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CheckBox chkQuad9 
      Caption         =   "Quad9"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   3360
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtZEN 
      Height          =   285
      Left            =   5880
      TabIndex        =   33
      ToolTipText     =   "Use this DNS server instead of DNS from network config"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CheckBox chkZenRBL 
      Caption         =   "Enable ZEN RBL"
      Height          =   495
      Left            =   3600
      TabIndex        =   31
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CheckBox chkAlienNIDS 
      Caption         =   "AlienVault OTX NIDS"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CheckBox chkAlienpDNS 
      Caption         =   "AlienVault OTX passive DNS"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   2400
      Value           =   1  'Checked
      Width           =   3015
   End
   Begin VB.CheckBox chkAlienVault 
      Caption         =   "AlienVault OTX API"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      ToolTipText     =   "Use AlienVault OTX API key (not required but provides greater number of lookup queries)"
      Top             =   1920
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CheckBox chkVirusTotal 
      Caption         =   "Disable VirusTotal.com API"
      Height          =   495
      Left            =   2640
      TabIndex        =   5
      Top             =   360
      Width           =   2295
   End
   Begin VB.CheckBox chkExcel 
      Caption         =   "Use Excel"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      ToolTipText     =   "Default output is CSV"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CheckBox chkReadCache 
      Caption         =   "Disable cache read"
      Height          =   195
      Left            =   360
      TabIndex        =   1
      ToolTipText     =   "Caching can limit API calls"
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox chkWriteCache 
      Caption         =   "Disable cache write"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txtTimeDelay 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Text            =   "15000"
      ToolTipText     =   "Default is 15000 miliseconds to do 4 lookups per minute. Do not adjust lower unless selected vendors allow more lookup queries"
      Top             =   480
      Width           =   1335
   End
   Begin VB.CheckBox chkCbR 
      Caption         =   "Carbon Black EDR API"
      Height          =   495
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Formerly called Cb Response"
      Top             =   3000
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CheckBox chkThreatGRID 
      Caption         =   "ThreatGRID"
      Height          =   495
      Left            =   360
      TabIndex        =   21
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox ChkSleep 
      Caption         =   "Sleep On Cached Lookup"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      ToolTipText     =   $"VTTL_GUI.frx":7AE6
      Top             =   720
      Width           =   2415
   End
   Begin VB.Frame FrameHashDomain 
      Height          =   5295
      Left            =   240
      TabIndex        =   48
      Top             =   1440
      Width           =   7575
      Begin VB.CheckBox chkSeclytics 
         Caption         =   "Seclytics"
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CheckBox chkPT 
         Caption         =   "RiskIQ"
         Height          =   495
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Previously PassiveTotal"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkET 
         Caption         =   "ET Intelligence"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Emerging Threats Intelligence from Proofpoint"
         Top             =   3000
         Width           =   1455
      End
      Begin VB.CheckBox chkThreatCrowd 
         Caption         =   "ThreatCrowd"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "No API key (limit queries to 4 per minute to avoid getting banned)"
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.OptionButton optionVTcategory 
         Caption         =   "Communicating"
         Height          =   255
         Index           =   2
         Left            =   5640
         TabIndex        =   54
         ToolTipText     =   "File communicating with queried domain/IP"
         Top             =   3720
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optionVTcategory 
         Caption         =   "Referrer"
         Height          =   255
         Index           =   1
         Left            =   4680
         TabIndex        =   53
         ToolTipText     =   "Files referencing queried domain/IP"
         Top             =   3720
         Width           =   975
      End
      Begin VB.OptionButton optionVTcategory 
         Caption         =   "Downloaded"
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   52
         ToolTipText     =   "Files download from queried domain/IP"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label chkVendor 
         Caption         =   "Display Results for"
         Height          =   735
         Left            =   3480
         TabIndex        =   60
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblVendorThreshold 
         Caption         =   "Vendor detection threshold:"
         Height          =   375
         Left            =   3360
         TabIndex        =   57
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label lblSampleCount 
         Caption         =   "Number of samples to report on:"
         Height          =   255
         Left            =   3360
         TabIndex        =   56
         Top             =   4080
         Width           =   3135
      End
      Begin VB.Label lblVTCategory 
         Caption         =   "VirusTotal Network Category Reporting:"
         Height          =   375
         Left            =   3360
         TabIndex        =   55
         Top             =   3480
         Width           =   3015
      End
      Begin VB.Label lblDNS 
         Caption         =   "DNS Server IP Address"
         Height          =   255
         Left            =   5640
         TabIndex        =   51
         ToolTipText     =   "DNS servers are optional. You may need to overide the DNS server for certain DBLs to work."
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.TextBox listscan 
      Height          =   285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cache"
      Height          =   1215
      Left            =   240
      TabIndex        =   68
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label lblTimeBetweenLookups 
      Caption         =   "milliseconds between lookups:"
      Height          =   495
      Left            =   5400
      TabIndex        =   47
      ToolTipText     =   "Pause between lookups"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'file IO
Const forwriting = 2
Const ForAppending = 8
Const ForReading = 1


Public paths As String 'used to identify external executables


'INI file declarations

Private Declare Function GetPrivateProfileString Lib "kernel32" _
   Alias "GetPrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As String, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long, _
   ByVal lpFileName As String) As Long
   
Private Declare Function WritePrivateProfileString Lib "kernel32" _
   Alias "WritePrivateProfileStringA" _
  (ByVal lpSectionName As String, _
   ByVal lpKeyName As String, _
   ByVal lpValue As String, _
   ByVal lpFileName As String) As Long
   


'End INI file declarations


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim intFrameHashDomainTop 'GUI tracking
Dim intPreviousTimeDelay 'GUI tracking
Public boolSaveConfigPrompt 'GUI tracking
Dim intSampleCount 'GUI tracking
Public boolInitialInstall

'whois
Private OTXwhocheck As Integer
Private arinValWhois As Integer
Private ripeValWhois As Integer
Private sysinternalsValWhois As Integer
Private NirSoftValWhois As Integer

Private freeGeoIPVal As Integer

'Intel feeds
Public WatchIntelURL As Boolean
Public staticIntel As Boolean
Public intelMulti As Boolean
Public intelMalware As Boolean
Public intelAttacker As Boolean
Public intelProxy As Boolean
Public ageFilter As Integer

Public dictCSVFeed

Public Property Let otxWhois(ByVal newvalue As Integer)
    OTXwhocheck = newvalue
End Property
Public Property Get otxWhois() As Integer
    otxWhois = OTXwhocheck
End Property

Public Property Let freeGeoIP(ByVal newvalue As Integer)
    freeGeoIPVal = newvalue
End Property
Public Property Get freeGeoIP() As Integer
    freeGeoIP = freeGeoIPVal
End Property


Public Property Let arinWhois(ByVal newvalue As Integer)
    arinValWhois = newvalue
End Property
Public Property Get arinWhois() As Integer
    arinWhois = arinValWhois
End Property
Public Property Let ripeWhois(ByVal newvalue As Integer)
    ripeValWhois = newvalue
End Property
Public Property Get ripeWhois() As Integer
    ripeWhois = ripeValWhois
End Property
Public Property Let sysinternalsWhois(ByVal newvalue As Integer)
    sysinternalsValWhois = newvalue
End Property
Public Property Get sysinternalsWhois() As Integer
    sysinternalsWhois = sysinternalsValWhois
End Property
Public Property Let NirSoftWhois(ByVal newvalue As Integer)
    NirSoftValWhois = newvalue
End Property
Public Property Get NirSoftWhois() As Integer
    NirSoftWhois = NirSoftValWhois
End Property
'end whois

Sub WriteINIvalue(strIniSectionF, strIniKeyF, StrIniValue)

m_lLastReturnCode = WritePrivateProfileString(strIniSectionF, strIniKeyF, StrIniValue, App.Path & "\vttl.ini")

End Sub
Function boolStringtoInt(stringBoolean)
If LCase(stringBoolean) = "true" Then
    boolStringtoInt = 1
Else
    boolStringtoInt = 0
End If
End Function
Function intToBool(stringInteger)
If IsNumeric(stringInteger) = True Then
    If stringInteger = 0 Then
        intToBool = "False"
    Else
        intToBool = "True"
    End If
End If
End Function

Sub ReadINIsettings()
strConfigPath = CreateFolder(App.Path & "\Config")
strIniPath = App.Path & "\vttl.ini"
strIniSection = "main"
strIniKey = "time_between_lookups"
Form1.txtTimeDelay.Text = readINIvalue(strIniSection, strIniPath, strIniKey) 'Restrict timeframe of event instances

strIniKey = "disable_CacheWrite"
tmpBool = readINIvalue(strIniSection, strIniPath, strIniKey)
Form1.chkWriteCache.Value = boolStringtoInt(tmpBool)
strIniKey = "disable_CacheRead"
Form1.chkReadCache.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "whoisCache"
Form1.chkWhoIsCache.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "cacheDomain"
chkDomainCache.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "deepIOCmatch"
Form1.chkDeepIOC.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_Excel"
Form1.chkExcel.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "SleepOnCachedLookup"
Form1.ChkSleep.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "InitialInstall"
boolInitialInstall = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "FeedAgeLimit"
tmpAgeFilter = readINIvalue(strIniSection, strIniPath, strIniKey)
If IsNumeric(tmpAgeFilter) Then
    ageFilter = tmpAgeFilter
End If

strIniSection = "vendor" ' VENDOR Section
strIniKey = "disable_VirusTotal"
Form1.chkVirusTotal.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "UseARIN"
arinWhois = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "UseRIPE"
ripeWhois = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "SysinternalsWhois"
sysinternalsWhois = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "NirSoft_WhosIP"
NirSoftWhois = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "useFreeGeoIP"
freeGeoIP = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_TIA"
Form1.chkTIA.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_CarbonBlack"
Form1.chkCbR.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_CarbonBlackEnterprise"
Form1.chkCBC.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_ThreatGRID"
Form1.chkThreatGRID.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "UseETIntelligence"
Form1.chkET.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))


strIniKey = "enable_ThreatCrowd"
Form1.chkThreatCrowd.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_ET"
Form1.chkET.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))


strIniKey = "enable_BlockLists"
Form1.chkBLchecks.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "disable_DomainBlockLists"
Form1.chkDomainBLchecks.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_ZEN"
Form1.chkZenRBL.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_URIBL"
Form1.chkURIBL.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_CBL"
Form1.chkCBL.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_ZDBL"
Form1.chkZDBL.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_Barracuda"
Form1.chkBarracuda.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_SURBL"
Form1.chkSURBL.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_SORBS"
Form1.chkSORBS.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_Quad9"
Form1.chkQuad9.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_CIF"
Form1.chkCIF.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_MalShare"
Form1.chkMalshare.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "useSeclytics"
Form1.chkSeclytics.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "EnablePassiveTotal"
Form1.chkPT.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "usePulsedive"
Form1.chkPulsedive.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_Shodan"
Form1.chkShodan.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "enable_AlienVault"
Form1.chkAlienVault.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "WatchIntelURLs"
WatchIntelURL = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "StaticIntel"
staticIntel = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "MultiFeed"
intelMulti = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "MalwareFeed"
intelMalware = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "AttackerFeed"
intelAttacker = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

strIniKey = "ProxyFeed"
intelProxy = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

'alienvault
strIniSection = "vendor_AlienVault"
strIniKey = "disable_whois"
'invert so we can say in the text enable instead of disable
otxWhois = invertDigit(boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey)))

strIniKey = "enable_passiveDNS"
Form1.chkAlienpDNS = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))
strIniKey = "enable_NIDS"
Form1.chkAlienNIDS = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))


'DNS Servers
strIniSection = "DNS_Server"
strIniKey = "Barracuda"
Form1.txtBarracuda.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "zen"
Form1.txtZEN.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "uribl"
Form1.txtURIBL.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "surbl"
Form1.txtSURBL.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "abuseat"
Form1.txtCBL.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "SORBS"
Form1.txtSORBS.Text = readINIvalue(strIniSection, strIniPath, strIniKey)

'VirusTotal
strIniSection = "VirusTotal"
strIniKey = "WebSamplesToCheck"
Form1.txtSampleCount = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "WebSampleCategory"
intTmpCategory = readINIvalue(strIniSection, strIniPath, strIniKey)
Select Case intTmpCategory
    Case 0
        optionVTcategory(0).Value = True
    Case 1
        optionVTcategory(1).Value = True
    Case 2
        optionVTcategory(2).Value = True
    Case Else
        optionVTcategory(2).Value = True
End Select
strIniKey = "WebSamplePositiveThreshold"
txtDetectThreshold.Text = readINIvalue(strIniSection, strIniPath, strIniKey)
strIniKey = "DisplayVendor"
StrIniValue = readINIvalue(strIniSection, strIniPath, strIniKey)
If StrIniValue <> "Vendor Name" And StrIniValue <> "" Then
    Form1.ComboVendor = StrIniValue
End If
strIniKey = "UseRegexForURL"
Form1.ChkURLregex.Value = boolStringtoInt(readINIvalue(strIniSection, strIniPath, strIniKey))

dictCSVFeed.RemoveAll
LoadCustomValDict App.Path & "\config\csvfeed.dat", dictCSVFeed 'load feed csv pointers


End Sub

Function invertDigit(bitFlip)
If bitFlip = 1 Then
    bitFlip = 0
Else
    bitFlip = 1
End If
invertDigit = bitFlip
End Function


Function readINIvalue(strIniSectionF, strIniPathF, strIniKeyF)
Dim strIniPath
Dim strIniSection
Dim strIniKey

Dim buf As String * 256
Dim Length As Long

    Length = GetPrivateProfileString( _
        strIniSectionF, strIniKeyF, "", _
        buf, Len(buf), strIniPathF)
    readINIvalue = Left$(buf, Length)

End Function



Private Sub btnEditDomain_Click()

Shell "notepad" & " " & Chr(34) & App.Path & "\VTTL_NoSubmit.txt" & Chr(34), vbNormalFocus

End Sub

Private Sub btnEditExact_Click()


Shell "notepad" & " " & Chr(34) & App.Path & "\VTTL_domains.txt" & Chr(34), vbNormalFocus

End Sub

Private Sub btnFolder_Click()
strReportsPath = App.Path & "\Reports"
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strReportsPath) = False Then
    objFSO.CreateFolder strReportsPath
End If
 Set WshShell = CreateObject("WScript.Shell")
    WshShell.CurrentDirectory = App.Path
    WshShell.Run "explorer.exe /e, " & App.Path & "\Reports"
    Set WshShell = Nothing
End Sub

Private Sub btnScanList_Click()
Dim vtlistLocation
vtlistLocation = App.Path & "\vtlist.txt"
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(vtlistLocation) = False Then
    cmdEditScanList.BackColor = &H8000000D
   
    MsgBox "The vtlist.txt file does not exist. Opening in notepad for you to save the list. To edit next time click the Edit Scan List button."
    Shell "notepad" & " " & Chr(34) & App.Path & "\vtlist.txt" & Chr(34), vbNormalFocus
    cmdEditScanList.BackColor = &H8000000F 'default customer
Sleep (100)
    Exit Sub
End If
If FrameHashDomain.Top = intFrameHashDomainTop Then
    FrameHashDomain.Top = 0
    FrameHashDomain.Height = FrameHashDomain.Height + 1445
    listscan.ZOrder 0
    listscan.Height = 6300
    listscan.Visible = True
    
    Set objTextFile = objFSO.OpenTextFile _
    (vtlistLocation, ForReading)
    If objTextFile.AtEndOfStream = False Then
        Form1.listscan.Text = objTextFile.ReadAll
    End If
    objTextFile.Close
    If CheckVTTLisHash = True Then
        Form1.optHash(0).Value = True
    Else
        Form1.optHash(1).Value = True
    End If
Else
    FrameHashDomain.Top = intFrameHashDomainTop
    FrameHashDomain.Height = 5295
    listscan.Visible = False
'If objFSO.FileExists(vtlistLocation) Then
'    objFSO.DeleteFile vtlistLocation, True
'End If
'Set objFile = objFSO.CreateTextFile(vtlistLocation, True)
'objFile.Write Form1.listscan.Text
'objFile.Close
End If
End Sub





Private Sub chkAlienNIDS_Click()

If chkAlienNIDS.Value = 0 Then Exit Sub


Set fso = CreateObject("Scripting.FileSystemObject")
If Form1.chkAlienVault = 0 Then
    tmpAnswer = MsgBox("AlienVault has been disabled. Would you like to enable?", vbYesNo, "Question")
    Form1.chkAlienVault = 1
    boolSaveConfigPrompt = True
ElseIf chkAlienNIDS.ToolTipText <> "" And chkAlienNIDS.Value = 1 Then
    MsgBox "An API key for AlienVault OTX has not been provided. An API key will be required to enable queries for AlienVault NIDS"
    boolSaveConfigPrompt = True
End If
AlienNIDS_Check 'sets chkAlienNIDS values

End Sub

Private Sub chkAlienVault_Click()
boolSaveConfigPrompt = True

End Sub

Private Sub chkBarracuda_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkBLchecks_Click()
boolSaveConfigPrompt = True
If chkBLchecks.Value = 1 Then
    setDBLEnabled True
Else
    setDBLEnabled False
End If
End Sub

Sub setDBLEnabled(boolRBL)
    Form1.chkCBL.Enabled = boolRBL
    chkZenRBL.Enabled = boolRBL
    chkSURBL.Enabled = boolRBL
    chkSORBS.Enabled = boolRBL
    chkBarracuda.Enabled = boolRBL
    Form1.chkZDBL.Enabled = boolRBL
    Form1.chkURIBL.Enabled = boolRBL
    txtZEN.Enabled = boolRBL
    txtURIBL.Enabled = boolRBL
    txtCBL.Enabled = boolRBL
    txtZenDBL.Enabled = boolRBL
    txtSURBL.Enabled = boolRBL
    txtSORBS.Enabled = boolRBL
    txtBarracuda.Enabled = boolRBL
End Sub

Private Sub chkCBC_Click()
If chkCBC.Value = 1 Then
strIniPath = App.Path & "\vttl.ini"
strIniSection = "vendor"
strIniKey = "CarbonBlackOrgKey"
CarbonBlackOrgKey = Form1.readINIvalue(strIniSection, strIniPath, strIniKey) 'Restrict timeframe of event instances
If CarbonBlackOrgKey = " " Or CarbonBlackOrgKey = "" Then
 newOrgKey = InputBox("Please provide your Carbon Black Org Key:", "VTTL INI Settings - " & App.Path)
 If newOrgKey <> " " And newOrgKey <> "" Then
    Form1.WriteINIvalue strIniSection, strIniKey, newOrgKey
 End If
End If

End If
boolSaveConfigPrompt = True
End Sub

Private Sub chkCBL_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkCbR_Click()
boolSaveConfigPrompt = True
HandleDisableFile Form1.chkCbR, App.Path & "\cb.disable", "Cb Response API"
End Sub

Private Sub chkCIF_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkDeepIOC_Click()
boolSaveConfigPrompt = True

End Sub

Private Sub chkDomainBLchecks_Click()
boolSaveConfigPrompt = True
If chkDomainBLchecks.Value = 1 Then
Form1.chkZDBL.Value = 0
Form1.chkZDBL.Enabled = False
Form1.chkURIBL.Value = 0
Form1.chkURIBL.Enabled = False
Else
Form1.chkZDBL.Enabled = True
Form1.chkURIBL.Enabled = True
End If
End Sub

Private Sub chkDomainCache_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkET_Click()
boolSaveConfigPrompt = True
HandleDisableFile Form1.chkCbR, App.Path & "\pp.disable", "Proofpoint ET Intelligence API"
End Sub

Private Sub chkExcel_Click()
boolSaveConfigPrompt = True
If chkExcel.Value = 1 Then
frmLaunch.txtFileName.Enabled = False
frmLaunch.txtFileName.ToolTipText = "Disabled when using Excel"
Else
frmLaunch.txtFileName.Enabled = True
End If
End Sub

Private Sub chkMalshare_Click()
boolSaveConfigPrompt = True
HandleDisableFile Form1.chkMalshare, App.Path & "\mals.disable", "MalShare API"
End Sub

Private Sub chkPT_Click()
HandleDisableFile Form1.chkPT, App.Path & "\pt.disable", "RiskIQ (PassiveTotal) API"
boolSaveConfigPrompt = True
End Sub

Private Sub chkPulsedive_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkQuad9_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkReadCache_Click()
boolSaveConfigPrompt = True
End Sub


Private Sub chkSeclytics_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkShodan_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub ChkSleep_Click()
boolSaveConfigPrompt = True
If ChkSleep.Value = 0 Then
    Form1.chkMalshare.Value = 0
    Form1.chkThreatCrowd.Value = 0
    
End If
End Sub

Private Sub chkSORBS_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkSURBL_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkThreatCrowd_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkThreatGRID_Click()
boolSaveConfigPrompt = True
HandleDisableFile Form1.chkThreatGRID, App.Path & "\tg.disable", "ThreatGRID API"
End Sub

Private Sub chkTIA_Click()
boolSaveConfigPrompt = True
HandleDisableFile Form1.chkTIA, App.Path & "\tia.disable", "ThreatIntelligenceAggregator API"
End Sub

Private Sub chkURIBL_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkVirusTotal_Click()
boolSaveConfigPrompt = True
If chkVirusTotal.Value = 1 Then
    txtSampleCount.Text = 0
    txtSampleCount.Enabled = False
    txtDetectThreshold.Enabled = False
    optionVTcategory(0).Enabled = False
    optionVTcategory(1).Enabled = False
    optionVTcategory(2).Enabled = False
    ComboVendor.Enabled = False
    If frmWhois.Visible = True Then frmWhois.chkVTwhois.Value = 0
Else
    txtSampleCount.Enabled = True
    txtDetectThreshold.Enabled = True
    optionVTcategory(0).Enabled = True
    optionVTcategory(1).Enabled = True
    optionVTcategory(2).Enabled = True
    ComboVendor.Enabled = True
    If frmWhois.Visible = True Then frmWhois.chkVTwhois.Value = 1
End If
    
End Sub

Private Sub chkWhoIsCache_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkWriteCache_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkZDBL_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub chkZenRBL_Click()
boolSaveConfigPrompt = True
End Sub

Private Sub cmdDictEdit_Click()
frmDict.Visible = True
frmDict.WindowState = vbNormal
End Sub

Private Sub cmdListEdit_Click()
'URLwatchlist.txt
'DNwatchlist.txt
End Sub

Private Sub cmdEditScanList_Click()
    Shell "notepad" & " " & Chr(34) & App.Path & "\vtlist.txt" & Chr(34), vbNormalFocus
End Sub

Private Sub cmdReload_Click()
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & "\vttl.ini") = True Then
    ReadINIsettings 'load settings
End If
End Sub

Private Sub Command1_Click()
If boolSaveConfigPrompt = True Then
    a = MsgBox("Do you want to save your config changes before proceeding?", vbYesNo, "Question")
    If a = vbYes Then
        saveBtn_Click
    End If
End If
If CheckVTTLisHash = True Then
    Form1.optHash(0).Value = True

Else
    Form1.optHash(1).Value = True
End If

frmLaunch.SetFormEnables True
frmLaunch.Visible = True
frmLaunch.WindowState = vbNormal

frmLaunch.SetFocus

End Sub



Sub AlienNIDS_Check()
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(App.Path & "\av.dat") = True And _
 fso.FileExists(App.Path & "\av.disable") = False Then
    Form1.chkAlienNIDS.Value = 1
    Form1.chkAlienNIDS.ToolTipText = ""
ElseIf fso.FileExists(App.Path & "\av.disable") = False Then
    'Form1.chkAlienNIDS.Value = 0
    Form1.chkAlienNIDS.ToolTipText = "Requires AlienVault OTX API key and paid subscription"
End If

End Sub


Sub LoadCustomValDict(strListPath, dictToLoad)
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(strListPath) Then
  Set objFile = objFSO.OpenTextFile(strListPath)
  Do While Not objFile.AtEndOfStream
    If Not objFile.AtEndOfStream Then 'read file
        On Error Resume Next
        strTmpData = objFile.ReadLine
        If InStr(strTmpData, "|") Then
          strTmpArrayDDNS = Split(strTmpData, "|")
          valueText = Mid(strTmpData, Len(strTmpArrayDDNS(0)) + 2)
          If dictToLoad.Exists(LCase(strTmpArrayDDNS(0))) = False Then _
          dictToLoad.Add LCase(strTmpArrayDDNS(0)), valueText
        Else
          If dictToLoad.Exists(LCase(strTmpData)) = False Then _
          dictToLoad.Add LCase(strTmpData), ""
        End If
        On Error GoTo 0
    End If
  Loop
End If
End Sub


Private Sub Form_Load()
'check last import time of geoip. Suggest importing a newer version.
'looking up IP addresses and no vendors select who return that data then suggest freegeoip. If low lookup threshold suggest adding sqlite if needed and importing geoip
'looking up domain/ip and no vendors select who return whois data then suggest using open whois lookups and external tools. Check lookup delay and don't suggest if too low.
'boolean setting to enable/disable suggestions
'support prompting for API keys in the GUI
'support tria.ge API

Set dictCSVFeed = CreateObject("Scripting.Dictionary")
intFrameHashDomainTop = 1440
Set fso = CreateObject("Scripting.FileSystemObject")
Form1.Caption = "VTTL INI Settings - " & App.Path
If fso.FileExists(App.Path & "\cb.dat") = True And _
 fso.FileExists(App.Path & "\cb.disable") = False Then
    Form1.chkCbR.Value = 1
End If
AlienNIDS_Check 'sets chkAlienNIDS values
If fso.FileExists(App.Path & "\tia.dat") = True And _
 fso.FileExists(App.Path & "\tia.disable") = False Then
    Form1.chkTIA.Value = 0
End If
If fso.FileExists(App.Path & "\tg.dat") = True And _
 fso.FileExists(App.Path & "\tg.disable") = False Then
    Form1.chkThreatGRID.Value = 1
End If
If fso.FileExists(App.Path & "\malshare.dat") = True And _
 fso.FileExists(App.Path & "\malshare.disable") = False Then
    Form1.chkMalshare.Value = 1
End If
If fso.FileExists(App.Path & "\vttl.ini") = True Then
    ReadINIsettings 'load settings
End If

If CheckVTTLisHash = True Then
    Form1.optHash(0).Value = True
Else
    Form1.optHash(1).Value = True
End If
If IsNumeric(txtTimeDelay.Text) Then
    intPreviousTimeDelay = txtTimeDelay.Text
Else
    intPreviousTimeDelay = 15000
End If

ReadVendorlist 'load vedor list
boolSaveConfigPrompt = False
queryExcel
paths = Environ$("path")
Unload frmWhois 'WORKAROUND: staring the app seems to always load frmWhois, which causes it to not have config due to variables not being loaded yet.


If fso.FileExists(App.Path & "\config\feedlist.dat") = False And fso.FileExists(App.Path & "\config\feedlist.default") = True Then
    fso.CopyFile App.Path & "\config\feedlist.default", App.Path & "\config\feedlist.dat"
End If
FrmFeed.ReadFeedlist "", "", False, False, False 'use "" parameter to load dict

'backwards compatibility update file path location
strConfigPath = CreateFolder(App.Path & "\Config")
compatibleConfigPath strConfigPath, "VTTL_NoSubmit.txt"
compatibleConfigPath strConfigPath, "VTTL_domains.txt"
compatibleConfigPath strConfigPath, "cc.dat"
compatibleConfigPath strConfigPath, "malhash.dat"
compatibleConfigPath strConfigPath, "whitehash.dat"
compatibleConfigPath strConfigPath, "IPDwatchlist.txt"
compatibleConfigPath strConfigPath, "\generics.alias"
compatibleConfigPath strConfigPath, "DNwatchlist.txt"
compatibleConfigPath strConfigPath, "KWordwatchlist.txt"
compatibleConfigPath strConfigPath, "URLwatchlist.txt"
'backwards compatibility update file path location

End Sub


Function CreateFolder(strFolderPath)
Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FolderExists(strFolderPath) = False Then _
objFSO.CreateFolder (strFolderPath)
CreateFolder = strFolderPath
End Function
Function compatibleConfigPath(strConfigFolderPath, strConfigFileName) 'moved config files out of current directory. This is backwards compatibility code.
  Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
  If objFSO.FileExists(strConfigFolderPath & "\" & strConfigFileName) Then
    compatibleConfigPath = strConfigFolderPath & "\" & strConfigFileName
  ElseIf objFSO.FileExists(App.Path & "\" & strConfigFileName) Then
    On Error Resume Next
    objFSO.MoveFile App.Path & "\" & strConfigFileName, strConfigFolderPath & "\"
    If Err.Number <> 0 Then
      compatibleConfigPath = App.Path & "\" & strConfigFileName
      MsgBox Err.Description
    Else
      compatibleConfigPath = strConfigFolderPath & "\" & strConfigFileName
    End If
    On Error GoTo 0
  Else
    compatibleConfigPath = strConfigFolderPath & "\" & strConfigFileName
  End If
End Function

Function FormCount(ByVal frmName As String) As Long
Dim frm As Form
For Each frm In Forms
    If StrComp(frm.Name, frmName, vbTextCompare) = 0 Then
        FormCount = FormCount + 1
    End If
Next
End Function

Private Sub Form_Unload(Cancel As Integer)
If FormCount("frmLaunch") = 1 Then Unload frmLaunch
If FormCount("frmDict") = 1 Then Unload frmDict
If FormCount("frmWhois") = 1 Then Unload frmWhois
If FormCount("frmFeed") = 1 Then Unload FrmFeed
End Sub
Sub HandleVisibility(formObject, boolVisible)

formObject.Visible = boolVisible
formObject.TabStop = boolVisible
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblWhois_Click()
frmWhois.Visible = True
frmWhois.ZOrder 0
End Sub

Private Sub lblWhois_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    lblWhois_Click
End If
End Sub

Private Sub optHash_Click(Index As Integer)
If optHash(1).Value = True Then
Form1.txtBarracuda.TabStop = False
'Form1.txtCBL.Visible = True
HandleVisibility Form1.txtCBL, True
HandleVisibility Form1.txtBarracuda, True
HandleVisibility Form1.txtSORBS, True
HandleVisibility Form1.txtURIBL, True
HandleVisibility Form1.txtZEN, True
HandleVisibility Form1.txtZenDBL, True
HandleVisibility Form1.txtSURBL, True
HandleVisibility Form1.chkAlienNIDS, True
HandleVisibility Form1.chkAlienpDNS, True
HandleVisibility lblWhois, True
HandleVisibility Form1.chkBarracuda, True
HandleVisibility Form1.chkBLchecks, True
HandleVisibility Form1.chkCBL, True
HandleVisibility Form1.chkDomainBLchecks, True
HandleVisibility Form1.chkQuad9, True
HandleVisibility Form1.chkSORBS, True
HandleVisibility Form1.chkSURBL, True
HandleVisibility Form1.chkURIBL, True
HandleVisibility Form1.chkZDBL, True
HandleVisibility Form1.chkZenRBL, True
lblDNS.Visible = True
chkPulsedive.Visible = True
chkShodan.Visible = True
lblSampleCount.Visible = True
lblVTCategory.Visible = True
HandleVisibility optionVTcategory(0), True
HandleVisibility optionVTcategory(1), True
HandleVisibility optionVTcategory(2), True
HandleVisibility txtSampleCount, True
HandleVisibility txtDetectThreshold, True
lblSampleCount.Visible = True
lblVendorThreshold.Visible = True
HandleVisibility ChkURLregex, True
HandleVisibility chkCIF, True
HandleVisibility chkQuad9, True
HandleVisibility Form1.chkET, False
HandleVisibility Form1.chkTIA, False
HandleVisibility Form1.chkMalshare, False
HandleVisibility Form1.chkPT, False
HandleVisibility Form1.chkCbR, False
HandleVisibility Form1.chkCBC, False
HandleVisibility Form1.ComboVendor, False
chkCBC.Visible = False
'HandleVisibility Form1.chkThreatCrowd, False
'Form1.chkET.Visible = False

If FormCount("frmLaunch") = 1 Then frmLaunch.SetFormEnables False 'form launch config
Else
HandleVisibility Form1.chkTIA, True
HandleVisibility Form1.chkMalshare, True
HandleVisibility Form1.chkPT, True
HandleVisibility Form1.chkCbR, True
HandleVisibility Form1.ComboVendor, True
HandleVisibility Form1.chkCBC, True
HandleVisibility Form1.chkThreatCrowd, True
HandleVisibility Form1.chkET, True
chkCBC.Visible = True
If FormCount("frmLaunch") = 1 Then frmLaunch.SetFormEnables False 'form launch config

lblDNS.Visible = False
chkPulsedive.Visible = False
chkShodan.Visible = False
HandleVisibility Form1.txtBarracuda, False
HandleVisibility Form1.txtCBL, False
HandleVisibility Form1.txtSORBS, False
HandleVisibility Form1.txtURIBL, False
HandleVisibility Form1.txtZEN, False
HandleVisibility Form1.txtZenDBL, False
HandleVisibility Form1.txtSURBL, False
HandleVisibility Form1.chkAlienNIDS, False
HandleVisibility Form1.chkAlienpDNS, False
HandleVisibility lblWhois, False
HandleVisibility Form1.chkBarracuda, False
HandleVisibility Form1.chkBLchecks, False
HandleVisibility Form1.chkCBL, False
HandleVisibility Form1.chkDomainBLchecks, False
HandleVisibility Form1.chkQuad9, False
HandleVisibility Form1.chkSORBS, False
HandleVisibility Form1.chkSURBL, False
HandleVisibility Form1.chkURIBL, False
HandleVisibility Form1.chkZDBL, False
HandleVisibility Form1.chkZenRBL, False
HandleVisibility txtSampleCount, False
lblSampleCount.Visible = False
lblVTCategory.Visible = False
HandleVisibility optionVTcategory(0), False
HandleVisibility optionVTcategory(1), False
HandleVisibility optionVTcategory(2), False
HandleVisibility txtSampleCount, False
HandleVisibility txtDetectThreshold, False
lblSampleCount.Visible = False
lblVendorThreshold.Visible = False
HandleVisibility ChkURLregex, False
HandleVisibility chkCIF, False
HandleVisibility chkQuad9, False
End If
End Sub

Private Sub saveBtn_Click()
strIniKey = "time_between_lookups"
StrIniValue = Form1.txtTimeDelay.Text 'Number of time to go back in logs for event aggregation:
IniSection = "main"
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "disable_CacheWrite"
StrIniValue = intToBool(Form1.chkWriteCache.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "disable_CacheRead"
StrIniValue = intToBool(Form1.chkReadCache.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "whoisCache"
StrIniValue = intToBool(Form1.chkWhoIsCache.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "CacheDomain"
StrIniValue = intToBool(Form1.chkDomainCache.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "SleepOnCachedLookup"
StrIniValue = intToBool(Form1.ChkSleep.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "enable_Excel"
StrIniValue = intToBool(Form1.chkExcel.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "deepIOCmatch"
StrIniValue = intToBool(Form1.chkDeepIOC.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "InitialInstall"
WriteINIvalue IniSection, strIniKey, boolInitialInstall

strIniKey = "FeedAgeLimit"
StrIniValue = ageFilter
WriteINIvalue IniSection, strIniKey, StrIniValue

'vendor
IniSection = "vendor"
strIniKey = "disable_VirusTotal"
StrIniValue = intToBool(Form1.chkVirusTotal.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_TIA"
StrIniValue = intToBool(Form1.chkTIA.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_AlienVault"
StrIniValue = intToBool(Form1.chkAlienVault.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_CarbonBlack"
StrIniValue = intToBool(Form1.chkCbR.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_CarbonBlackEnterprise"
StrIniValue = intToBool(Form1.chkCBC)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_ThreatGRID"
StrIniValue = intToBool(Form1.chkThreatGRID.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "UseETIntelligence"
StrIniValue = intToBool(Form1.chkET.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue 'vendor section
strIniKey = "enable_ZEN"
StrIniValue = intToBool(Form1.chkZenRBL.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_URIBL"
StrIniValue = intToBool(Form1.chkURIBL.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_CBL"
StrIniValue = intToBool(Form1.chkCBL.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_ZDBL"
StrIniValue = intToBool(Form1.chkZDBL.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_Barracuda"
StrIniValue = intToBool(Form1.chkBarracuda.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_SURBL"
StrIniValue = intToBool(Form1.chkSURBL.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_SORBS"
StrIniValue = intToBool(Form1.chkSORBS.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_Quad9"
StrIniValue = intToBool(Form1.chkQuad9.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_CIF"
StrIniValue = intToBool(Form1.chkCIF.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_AlienVault"
StrIniValue = intToBool(Form1.chkAlienVault.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue 'vendor section
strIniKey = "enable_MalShare"
StrIniValue = intToBool(Form1.chkMalshare.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_ThreatCrowd"
StrIniValue = intToBool(Form1.chkThreatCrowd.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_ET"
StrIniValue = intToBool(Form1.chkET.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "Disable_DomainBlockLists"
StrIniValue = intToBool(Form1.chkDomainBLchecks.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_BlockLists"
StrIniValue = intToBool(Form1.chkBLchecks.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "useSeclytics"
StrIniValue = intToBool(Form1.chkSeclytics.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "EnablePassiveTotal"
StrIniValue = intToBool(Form1.chkPT.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "usePulsedive"
StrIniValue = intToBool(Form1.chkPulsedive.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_Shodan"
StrIniValue = intToBool(Form1.chkShodan.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "UseARIN"
StrIniValue = intToBool(arinWhois)
WriteINIvalue IniSection, strIniKey, StrIniValue 'vendor section
strIniKey = "UseRIPE"
StrIniValue = intToBool(ripeWhois)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "SysinternalsWhois"
StrIniValue = intToBool(sysinternalsWhois)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "NirSoft_WhosIP"
StrIniValue = intToBool(NirSoftWhois)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "useFreeGeoIP"
StrIniValue = intToBool(freeGeoIP)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "WatchIntelURLs"
StrIniValue = intToBool(WatchIntelURL)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "StaticIntel"
StrIniValue = intToBool(staticIntel)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "MultiFeed"
StrIniValue = intToBool(intelMulti)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "MalwareFeed"
StrIniValue = intToBool(intelMalware)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "AttackerFeed"
StrIniValue = intToBool(intelAttacker)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "ProxyFeed"
StrIniValue = intToBool(intelProxy)
WriteINIvalue IniSection, strIniKey, StrIniValue


'AlienVault section
IniSection = "vendor_AlienVault"
strIniKey = "disable_whois"
StrIniValue = intToBool(invertDigit(otxWhois))
WriteINIvalue IniSection, strIniKey, StrIniValue


strIniKey = "enable_passiveDNS"
StrIniValue = intToBool(Form1.chkAlienpDNS.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "enable_NIDS"
StrIniValue = intToBool(Form1.chkAlienNIDS)
WriteINIvalue IniSection, strIniKey, StrIniValue
If Form1.chkAlienNIDS = 1 Then 'we are using an API key
    strIniKey = "use_AlienVaultAPIkey"
    StrIniValue = intToBool(Form1.chkAlienNIDS) 'use same value for as AlienNIDS as use_AlienVaultAPIkey
    WriteINIvalue IniSection, strIniKey, StrIniValue
End If
'DNS Server section
IniSection = "DNS_Server"
strIniKey = "Barracuda"
StrIniValue = Form1.txtBarracuda.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "zen"
StrIniValue = Form1.txtZEN.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "uribl"
StrIniValue = Form1.txtURIBL.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "surbl"
StrIniValue = Form1.txtSURBL.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "abuseat"
StrIniValue = Form1.txtCBL.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "SORBS"
StrIniValue = Form1.txtSORBS.Text
WriteINIvalue IniSection, strIniKey, StrIniValue


'VirusTotal Web Samples
IniSection = "VirusTotal"
strIniKey = "WebSamplesToCheck"
StrIniValue = Form1.txtSampleCount.Text
WriteINIvalue IniSection, strIniKey, StrIniValue
strIniKey = "WebSampleCategory"

For intArrayCount = 0 To 2
    If optionVTcategory(intArrayCount).Value = True Then
        StrIniValue = intArrayCount
    End If
Next
WriteINIvalue IniSection, strIniKey, StrIniValue
 
 strIniKey = "WebSamplePositiveThreshold"
StrIniValue = txtDetectThreshold.Text
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "UseRegexForURL"
StrIniValue = intToBool(Form1.ChkURLregex.Value)
WriteINIvalue IniSection, strIniKey, StrIniValue

strIniKey = "DisplayVendor"
StrIniValue = Form1.ComboVendor
If InStr(StrIniValue, " ") > 0 Then
  strVendorName = Left(StrIniValue, InStr(StrIniValue, " ") - 1)
End If
If StrIniValue <> "Vendor Name" And StrIniValue <> "" Then
    WriteINIvalue IniSection, strIniKey, StrIniValue
End If

'save feed list
FrmFeed.ReadFeedlist "", "", True, False, False

'save CSV feed column associations
csvDatPath = App.Path & "\config\csvfeed.dat"
Set fsoCSV = CreateObject("Scripting.FileSystemObject")
If fsoCSV.FileExists(csvDatPath) Then fsoCSV.DeleteFile (csvDatPath)
For Each feedName In dictCSVFeed
    LogData CStr(csvDatPath), feedName & "|" & dictCSVFeed.Item(feedName)
Next
boolSaveConfigPrompt = False
End Sub





Private Sub Timer_ButtonFlash_Timer()


Sleep (400)
End Sub

Private Sub Timer_FlashButton_Timer()
Timer_ButtonFlash.Enabled = False

'cmdEditScanList.BackColor = &H8000000F 'default customer


End Sub

Private Sub txtBarracuda_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtCBL_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtFeed_Click()
FrmFeed.Visible = True
FrmFeed.ZOrder 0
End Sub



Private Sub txtFeed_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    txtFeed_Click
End If
End Sub

Private Sub txtSampleCount_Change()
If IsNumeric(txtSampleCount.Text) Then
    If txtSampleCount.Text < 0 Then txtSampleCount.Text = 0
    intSampleCount = txtSampleCount.Text
    
        
    If intSampleCount = 0 Then
        txtDetectThreshold.Enabled = False
    Else
        txtDetectThreshold.Enabled = True
    End If
ElseIf txtSampleCount.Text = "" Then
    txtSampleCount.Text = 0

Else
    txtSampleCount.Text = intSampleCount
End If
End Sub

Private Sub txtSORBS_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtSURBL_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtTimeDelay_Change()
If IsNumeric(txtTimeDelay.Text) Then
    If txtTimeDelay.Text < 10000 Then
        Form1.chkThreatCrowd.Enabled = False
        Form1.chkMalshare.Enabled = False
    Else
        Form1.chkThreatCrowd.Enabled = True
        Form1.chkMalshare.Enabled = True
    End If
    intPreviousTimeDelay = txtTimeDelay.Text
Else
    txtTimeDelay.Text = intPreviousTimeDelay
End If
End Sub

Function CheckVTTLisHash()
Dim vttlLine As String
Set fsoIH = CreateObject("Scripting.FileSystemObject")

If fsoIH.FileExists(App.Path & "\vtlist.txt") = True Then
    Open App.Path & "\vtlist.txt" For Input As #1
        If Not EOF(1) Then
            Line Input #1, vttlLine

        End If
        
    Close #1
    If IsHash(vttlLine) Then
        CheckVTTLisHash = True
        Exit Function
    Else
        CheckVTTLisHash = False
    End If
Else
    CheckVTTLisHash = False
End If
End Function
Function IsHash(TestString)

    Dim sTemp
    Dim iLen
    Dim iCtr
    Dim sChar
    
    'returns true if all characters in a string are alphabetical
    '   or numeric
    'returns false otherwise or for empty string
    
    sTemp = TestString
    iLen = Len(sTemp)
    If iLen > 0 Then
        For iCtr = 1 To iLen
            sChar = Mid(sTemp, iCtr, 1)
            If IsNumeric(sChar) Or "a" = LCase(sChar) Or "b" = LCase(sChar) Or "c" = LCase(sChar) Or "d" = LCase(sChar) Or "e" = LCase(sChar) Or "f" = LCase(sChar) Then
              'allowed characters for hash (hex)
            Else
              IsHash = False
              Exit Function
            End If
        Next
    
    IsHash = True
    Else
      IsHash = False
    End If
    
End Function

Private Sub txtURIBL_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtZEN_Change()
boolSaveConfigPrompt = True
End Sub

Private Sub txtZenDBL_Change()
boolSaveConfigPrompt = True
End Sub

Sub ReadVendorlist()

Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
If objFSO.FileExists(compatibleConfigPath(App.Path & "\config", "VendorList.txt")) Then
    Set objFile = objFSO.OpenTextFile(compatibleConfigPath(App.Path & "\config", "VendorList.txt"))
    Do While Not objFile.AtEndOfStream
      If Not objFile.AtEndOfStream Then 'read file
          On Error Resume Next
          strVendorName = objFile.ReadLine

            ComboVendor.AddItem strVendorName
        
      End If
    Loop
  End If

End Sub

Public Function LogData(strLogName As String, LineData As String)
Const ForAppending = 8
Dim fs, f, strLogPath
If InStr(strLogName, "\") = 0 Then
    strLogPath = App.Path & "\" & strLogName
Else
    strLogPath = strLogName

End If
Set fs = CreateObject("Scripting.FileSystemObject")

'if no evidence is pending being gathered then roll logs.
'archive log files so we don't generate one single large file
If fs.FileExists(strLogPath) = False Then
      'Creates a replacement text file
      fs.CreateTextFile strLogPath, True
  Else
    Set objFile = fs.GetFile(strLogPath)
    If objFile.Size > 5000000 Then
      LogFileNameCounter = 0
          Logfilenamefree = False
          Do While Logfilenamefree = False
              LogFileNameCounter = LogFileNameCounter + 1
              OutputFile = strLogPath & LogFileNameCounter

              If fs.FileExists(OutputFile) <> True Then
                  Logfilenamefree = True
              End If

          Loop
          If fs.FileExists(OutputFile) = False Then
            On Error Resume Next
            fs.CopyFile strLogPath, OutputFile, True
             fs.DeleteFile strLogPath
             fs.CreateTextFile strLogPath, True
             On Error GoTo 0
          End If

    End If
  End If
'end archive log files so we don't generate one single large file

' Note: Set the last parameter to True, it will automatically create a new text file if it doesn't exist
On Error Resume Next

Set f = fs.OpenTextFile(strLogPath, ForAppending, True)
    f.WriteLine LineData
    f.Close
    
LogData = Err.Number

On Error GoTo 0

Set fs = Nothing
Set f = Nothing
End Function

'create or delete .disable file for each vendor. Just exits now since .disable are no longer used
Private Sub HandleDisableFile(formCheckObject, strDisableFilePath, strVendorName)
Exit Sub 'no longer use .disablefile
Set fso = CreateObject("Scripting.FileSystemObject")
If formCheckObject.Value = 0 And _
 fso.FileExists(strDisableFilePath) = False And _
 (boolInitialInstall = False Or fso.FileExists(Replace(LCase(strDatPath), ".disable", ".dat")) = True) Then 'if there is a dat file and we are disabling then create .disable file
    Set a = fso.CreateTextFile(strDisableFilePath, True)
    a.WriteLine (strVendorName & " has been disabled. Delete this file to enable")
    a.Close
ElseIf fso.FileExists(strDisableFilePath) = True Then
    On Error Resume Next
    fso.DeleteFile (strDisableFilePath)
    On Error GoTo 0
End If
End Sub

Sub queryExcel()
Dim oRegistry 'https://www.devhut.net/2013/10/23/determine-installed-version-of-any-ms-office-program-vbscript/
Dim oFSO
Dim sKey
Dim sAppExe
Dim sValue
Dim sAppVersion
Const HKEY_LOCAL_MACHINE = &H80000002
 
Set oRegistry = GetObject("winmgmts:{impersonationLevel=impersonate}//./root/default:StdRegProv")
Set oFSO = CreateObject("Scripting.FileSystemObject")
sKey = "Software\Microsoft\Windows\CurrentVersion\App Paths"
sAppExe = "excel.exe"

oRegistry.GetStringValue HKEY_LOCAL_MACHINE, sKey & "\" & sAppExe, "", sValue
If IsNull(sValue) = False Then
    strVersion = oFSO.GetFileVersion(sValue)
    If Len(strVersion) > 0 Then
        If InStr(strVersion, ".") > 0 Then
            If IsNumeric(Left(strVersion, 1)) = True Then
                chkExcel.Enabled = True
                chkExcel.ToolTipText = "Output to Excel instead of CSV"
            End If
        End If
    End If
Else
    Form1.chkExcel.Enabled = False
    chkExcel.ToolTipText = "Excel is not installed"
    chkExcel.Value = 0
End If
Set oFSO = Nothing
Set oRegistry = Nothing
End Sub
