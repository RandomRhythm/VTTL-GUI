VERSION 5.00
Begin VB.Form frmWhois 
   Caption         =   "VTTL - Whois"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3615
   Icon            =   "frmWhois.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6015
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "IP Geolocation"
      Height          =   855
      Left            =   600
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
      Begin VB.CheckBox chkFreeGeoIP 
         Caption         =   "FreeGeoIP"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CheckBox chkVTwhois 
      Caption         =   "VirusTotal"
      Enabled         =   0   'False
      Height          =   495
      Left            =   720
      TabIndex        =   8
      ToolTipText     =   "VirusTotal whois will be enabled when using VirusTotal"
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Domain Whois"
      Height          =   855
      Left            =   600
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
      Begin VB.CheckBox chkSysinternals 
         Caption         =   "Sysinternals Whois"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "Use external command line whois tool from Sysinternals"
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkWhoisAlien 
      Caption         =   "AlienVault OTX"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      ToolTipText     =   "AlienVault OTX whois API"
      Top             =   600
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Domain and IP"
      Height          =   1215
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Frame frameIP 
      Caption         =   "IP Whois"
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Width           =   2415
      Begin VB.CheckBox chkWhosIP 
         Caption         =   "NirSoft WhosIP"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Use external NirSoft whosip command line tool"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CheckBox chkRIPE 
         Caption         =   "RIPE"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Use RIPE web API"
         Top             =   720
         Width           =   1215
      End
      Begin VB.CheckBox chkARIN 
         Caption         =   "ARIN"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Use ARIN web API"
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmWhois"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkARIN_Click()
Form1.boolSaveConfigPrompt = True
Form1.arinWhois = chkARIN.Value
End Sub

Private Sub chkFreeGeoIP_Click()
Form1.boolSaveConfigPrompt = True
Form1.freeGeoIP = chkFreeGeoIP.Value
End Sub

Private Sub chkRIPE_Click()
Form1.boolSaveConfigPrompt = True
Form1.ripeWhois = chkRIPE.Value
End Sub

Private Sub chkSysinternals_Click()
'This check box  is only enabled when whois executable is in one of the folders from the path variable or same as exe or current directory

Form1.boolSaveConfigPrompt = True

If chkSysinternals.Value = 1 Then

    If pathEnum("whois.exe") = False Then
        chkSysinternals.Value = 0
        MsgBox ("The Sysinternals whois.exe execuable does not exist in the path. Please download it and put it in the path.")
    End If
End If

Form1.sysinternalsWhois = chkSysinternals.Value

End Sub

Private Sub chkWhoisAlien_Click()
Form1.boolSaveConfigPrompt = True
Form1.otxWhois = chkWhoisAlien.Value
End Sub

Private Sub chkWhosIP_Click()
Form1.boolSaveConfigPrompt = True
If chkWhosIP.Value = 1 Then

If pathEnum("whosip.exe") = False Then
    chkWhosIP.Value = 0
    MsgBox ("The NirSoft whosip.exe execuable does not exist in the path. Please download it and put it in the path.")
End If
End If
Form1.NirSoftWhois = frmWhois.chkWhosIP.Value
End Sub

Private Sub Form_Load()
chkWhoisAlien.Value = Form1.otxWhois
chkARIN.Value = Form1.arinWhois
chkRIPE.Value = Form1.ripeWhois
chkFreeGeoIP.Value = Form1.freeGeoIP
frmWhois.chkSysinternals = Form1.sysinternalsWhois
frmWhois.chkWhosIP = Form1.NirSoftWhois
chkVTwhois.Value = Form1.invertDigit(Form1.chkVirusTotal)

End Sub

Function pathEnum(strFile2Check)
boolPathExists = False
Set objectfso = CreateObject("Scripting.FileSystemObject")
If InStr(Form1.paths, ";") > 0 Then
    arrayPaths = Split(Form1.paths, ";")
    For Each folderPath In arrayPaths
        If objectfso.FileExists(folderPath & "\" & strFile2Check) Then
            boolPathExists = True
        End If
    Next
End If
If boolPathExists = False Then
    If objectfso.FileExists(App.Path & "\" & strFile2Check) = True Then
        boolPathExists = True
    End If
End If
pathEnum = boolPathExists
End Function
