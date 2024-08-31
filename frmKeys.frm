VERSION 5.00
Begin VB.Form frmKeys 
   Caption         =   "Form2"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11580
   LinkTopic       =   "Form2"
   ScaleHeight     =   5850
   ScaleWidth      =   11580
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox ComboAPI 
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   240
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "frmKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboAPI_Change()
InputBox ("would you like to save your change to " & previousText & "?")
End Sub

Private Sub Form_Load()
strRandom = "4bv3nT9vrkJpj3QyueTvYFBMIvMOllyuKy3d401Fxaho6DQTbPafyVmfk8wj1bXF" 'encryption key. Change if you want but can only decrypt with same key
frmKeys.ComboAPI.AddItem ("VirusTotal")
frmKeys.ComboAPI.AddItem ("AlienVault OTX")
frmKeys.ComboAPI.AddItem ("ThreatGRID")
frmKeys.ComboAPI.AddItem ("Emerging Threats ET Intelligence")
frmKeys.ComboAPI.AddItem ("Malshare")
frmKeys.ComboAPI.AddItem ("Carbon Black EDR/Hosted EDR (formally Cb Response)")
frmKeys.ComboAPI.AddItem ("Carbon Black Enterprise EDR (formally ThreatHunter)")
frmKeys.ComboAPI.AddItem ("ThreatGRID")
frmKeys.ComboAPI.AddItem ("ThreatIntelligenceAggregator (TIA)")
frmKeys.ComboAPI.AddItem ("RiskIQ")
frmKeys.ComboAPI.AddItem ("Collective Intelligence Framework (CIF)")
frmKeys.ComboAPI.AddItem ("SecLytics")
frmKeys.ComboAPI.AddItem ("Pulsedive")
strTmpData = Decrypt(strTmpData, strRandom)
End Sub

Sub loadAPIkey(strVendorText)


End Sub


Function encrypt(StrText, key)
  Dim lenKey, KeyPos, LenStr, x, Newstr
   
  Newstr = ""
  lenKey = Len(key)
  KeyPos = 1
  LenStr = Len(StrText)
  StrTmpText = StrReverse(StrText)
  For x = 1 To LenStr
       Newstr = Newstr & Chr(Asc(Mid(StrTmpText, x, 1)) + Asc(Mid(key, KeyPos, 1)))
       KeyPos = KeyPos + 1
       If KeyPos > lenKey Then KeyPos = 1
  Next
  encrypt = Newstr
 End Function

 
Function Decrypt(StrText, key)
  Dim lenKey, KeyPos, LenStr, x, Newstr
   
  Newstr = ""
  lenKey = Len(key)
  KeyPos = 1
  LenStr = Len(StrText)
   
  StrText = StrReverse(StrText)
  For x = LenStr To 1 Step -1
       On Error Resume Next
       Newstr = Newstr & Chr(Asc(Mid(StrText, x, 1)) - Asc(Mid(key, KeyPos, 1)))
       If Err.Number <> 0 Then
        MsgBox "error with char " & Chr(34) & Asc(Mid(StrText, x, 1)) - Asc(Mid(key, KeyPos, 1)) & Chr(34) & " At position " & KeyPos & vbCrLf & Mid(StrText, x, 1) & Mid(key, KeyPos, 1) & vbCrLf & Asc(Mid(StrText, x, 1)) & Asc(Mid(key, KeyPos, 1))
        wscript.quit (11)
       End If
       On Error GoTo 0
       KeyPos = KeyPos + 1
       If KeyPos > lenKey Then KeyPos = 1
       Next
       Newstr = StrReverse(Newstr)
       Decrypt = Newstr
End Function
