Attribute VB_Name = "VisualStyles"
Option Explicit
Public Declare Function ActivateVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByVal pszSubAppName As Long = 0, Optional ByVal pszSubIdList As Long = 0) As Long
Public Declare Function RemoveVisualStyles Lib "uxtheme" Alias "SetWindowTheme" (ByVal hWnd As Long, Optional ByRef pszSubAppName As String = " ", Optional ByRef pszSubIdList As String = " ") As Long
Public Declare Function GetVisualStyles Lib "uxtheme" Alias "GetWindowTheme" (ByVal hWnd As Long) As Long
Private Type TINITCOMMONCONTROLSEX
dwSize As Long
dwICC As Long
End Type
Private Type TRELEASE
IUnk As IUnknown
VTable(0 To 2) As Long
VTableHeaderPointer As Long
End Type
Private Type TRACKMOUSEEVENTSTRUCT
cbSize As Long
dwFlags As Long
hWndTrack As Long
dwHoverTime As Long
End Type
Private Enum UxThemeButtonParts
BP_PUSHBUTTON = 1
BP_RADIOBUTTON = 2
BP_CHECKBOX = 3
BP_GROUPBOX = 4
BP_USERBUTTON = 5
End Enum
Private Enum UxThemeButtonStates
PBS_NORMAL = 1
PBS_HOT = 2
PBS_PRESSED = 3
PBS_DISABLED = 4
PBS_DEFAULTED = 5
End Enum
Private Type POINTAPI
X As Long
Y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Type PAINTSTRUCT
hDC As Long
fErase As Long
RCPaint As RECT
fRestore As Long
fIncUpdate As Long
RGBReserved(0 To 31) As Byte
End Type
Private Type DLLVERSIONINFO
cbSize As Long
dwMajor As Long
dwMinor As Long
dwBuildNumber As Long
dwPlatformID As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function InitCommonControlsEx Lib "comctl32" (ByRef ICCEX As TINITCOMMONCONTROLSEX) As Long
Private Declare Function SetErrorMode Lib "kernel32" (ByVal wMode As Long) As Long
Private Declare Function DllGetVersion Lib "comctl32" (ByRef pdvi As DLLVERSIONINFO) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ExtSelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal fnMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateW" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lData As Long, ByVal wData As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal fFlags As Long) As Long
Private Declare Function GetProp Lib "user32" Alias "GetPropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function SetProp Lib "user32" Alias "SetPropW" (ByVal hWnd As Long, ByVal lpString As Long, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias "RemovePropW" (ByVal hWnd As Long, ByVal lpString As Long) As Long
Private Declare Function BeginPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32" (ByVal hWnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Any, ByVal bErase As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function TrackMouseEvent Lib "user32" (ByRef lpEventTrack As TRACKMOUSEEVENTSTRUCT) As Long
Private Declare Function TransparentBlt Lib "msimg32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, ByVal crTransparent As Long) As Long
Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme" (ByVal Theme As Long, iPartId As Long, iStateId As Long) As Long
Private Declare Function DrawThemeParentBackground Lib "uxtheme" (ByVal hWnd As Long, ByVal hDC As Long, ByRef pRect As RECT) As Long
Private Declare Function DrawThemeBackground Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef pClipRect As RECT) As Long
Private Declare Function DrawThemeText Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, ByRef pRect As RECT) As Long
Private Declare Function OpenThemeData Lib "uxtheme" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
Private Declare Function CloseThemeData Lib "uxtheme" (ByVal Theme As Long) As Long
Private Declare Function GetThemeBackgroundRegion Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pRect As RECT, ByRef hRgn As Long) As Long
Private Declare Function GetThemeBackgroundContentRect Lib "uxtheme" (ByVal Theme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByRef pBoundingRect As RECT, ByRef pContentRect As RECT) As Long
Private Declare Function IsAppThemed Lib "uxtheme" () As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Long
Private Declare Function GetThemeAppProperties Lib "uxtheme" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function SetWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Function RemoveWindowSubclass Lib "comctl32" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "comctl32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcW" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const ICC_STANDARD_CLASSES As Long = &H4000
Private Const STAP_ALLOW_CONTROLS As Long = (1 * (2 ^ 1))
Private Const S_OK As Long = &H0
Private Const UIS_CLEAR As Long = 2
Private Const UISF_HIDEFOCUS As Long = &H1
Private Const UISF_HIDEACCEL As Long = &H2
Private Const WM_CHANGEUISTATE As Long = &H127
Private Const WM_SETFOCUS As Long = &H7
Private Const WM_KILLFOCUS As Long = &H8
Private Const WM_ENABLE As Long = &HA
Private Const WM_SETREDRAW As Long = &HB
Private Const WM_PAINT As Long = &HF
Private Const WM_NCPAINT As Long = &H85
Private Const WM_NCDESTROY As Long = &H82
Private Const BM_GETSTATE As Long = &HF2
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEHOVER As Long = &H2A1
Private Const WM_MOUSELEAVE As Long = &H2A3
Private Const WM_PRINTCLIENT As Long = &H318
Private Const BST_PUSHED As Long = &H4
Private Const BST_FOCUS As Long = &H8
Private Const DT_CALCRECT As Long = &H400
Private Const DT_CENTER As Long = &H1
Private Const DT_WORDBREAK As Long = &H10
Private Const TME_HOVER As Long = 1
Private Const TME_LEAVE As Long = 2
Private Const RGN_DIFF As Long = 4
Private Const RGN_COPY As Long = 5
Private Const DST_ICON As Long = &H3
Private Const DST_BITMAP As Long = &H4
Private Const DSS_DISABLED As Long = &H20

Public Sub InitVisualStyles()
If App.LogMode <> 0 Then Call InitReleaseVisualStyles(AddressOf ReleaseVisualStyles)
Dim ICCEX As TINITCOMMONCONTROLSEX
With ICCEX
.dwSize = LenB(ICCEX)
.dwICC = ICC_STANDARD_CLASSES
End With
InitCommonControlsEx ICCEX
End Sub

Private Sub InitReleaseVisualStyles(ByVal Address As Long)
Static Release As TRELEASE
If Release.VTableHeaderPointer <> 0 Then Exit Sub
If GetComCtlVersion >= 6 Then
    Release.VTable(2) = Address
    Release.VTableHeaderPointer = VarPtr(Release.VTable(0))
    CopyMemory Release.IUnk, VarPtr(Release.VTableHeaderPointer), 4
End If
End Sub

Private Function ReleaseVisualStyles() As Long
Const SEM_NOGPFAULTERRORBOX As Long = &H2
SetErrorMode SEM_NOGPFAULTERRORBOX
End Function

Public Sub SetupVisualStyles(ByVal Form As VB.Form)
If GetComCtlVersion() >= 6 Then SendMessage Form.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS Or UISF_HIDEACCEL), ByVal 0&
If EnabledVisualStyles() = False Then Exit Sub
Dim CurrControl As Control
For Each CurrControl In Form.Controls
    Select Case TypeName(CurrControl)
        Case "Frame"
            SetWindowSubclass CurrControl.hWnd, AddressOf RedirectFrame, ObjPtr(CurrControl), 0
        Case "CommandButton", "CommandButtonW", "CheckBox", "CheckBoxW", "OptionButton", "OptionButtonW"
            If CurrControl.Style = vbButtonGraphical Then
                If CurrControl.Enabled = True Then SetProp CurrControl.hWnd, StrPtr("Enabled"), 1
                SetWindowSubclass CurrControl.hWnd, AddressOf RedirectButton, ObjPtr(CurrControl), ObjPtr(CurrControl)
            End If
    End Select
Next CurrControl
End Sub

Public Function EnabledVisualStyles() As Boolean
If GetComCtlVersion() >= 6 Then
    If IsThemeActive() <> 0 Then
        If IsAppThemed() <> 0 Then
            EnabledVisualStyles = True
        ElseIf (GetThemeAppProperties() And STAP_ALLOW_CONTROLS) <> 0 Then
            EnabledVisualStyles = True
        End If
    End If
End If
End Function

Public Function GetComCtlVersion() As Long
Static Done As Boolean, Value As Long
If Done = False Then
    Dim Version As DLLVERSIONINFO
    On Error Resume Next
    Version.cbSize = LenB(Version)
    If DllGetVersion(Version) = S_OK Then Value = Version.dwMajor
    Done = True
End If
GetComCtlVersion = Value
End Function

Private Function RedirectFrame(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Select Case wMsg
    Case WM_PRINTCLIENT, WM_MOUSELEAVE
        RedirectFrame = DefWindowProc(hWnd, wMsg, wParam, lParam)
        Exit Function
End Select
RedirectFrame = DefSubclassProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_NCDESTROY Then Call RemoveRedirectFrame(hWnd, uIdSubclass)
End Function

Private Sub RemoveRedirectFrame(ByVal hWnd As Long, ByVal uIdSubclass As Long)
RemoveWindowSubclass hWnd, AddressOf RedirectFrame, uIdSubclass
End Sub

Private Function RedirectButton(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal Button As Object) As Long
Dim SetRedraw As Boolean
Select Case wMsg
    Case WM_NCPAINT
        Exit Function
    Case WM_PAINT
        If IsWindowVisible(hWnd) <> 0 And GetVisualStyles(hWnd) <> 0 Then
            Dim PS As PAINTSTRUCT
            SetProp hWnd, StrPtr("Painted"), 1
            Call DrawButton(hWnd, BeginPaint(hWnd, PS), Button)
            EndPaint hWnd, PS
            Exit Function
        End If
    Case WM_SETFOCUS, WM_ENABLE
        If IsWindowVisible(hWnd) <> 0 Then
            SetRedraw = True
            SendMessage hWnd, WM_SETREDRAW, 0, ByVal 0&
        End If
End Select
RedirectButton = DefSubclassProc(hWnd, wMsg, wParam, lParam)
If wMsg = WM_NCDESTROY Then
    Call RemoveRedirectButton(hWnd, uIdSubclass)
    RemoveProp hWnd, StrPtr("Enabled"): RemoveProp hWnd, StrPtr("Hot"): RemoveProp hWnd, StrPtr("Painted")
ElseIf IsWindow(hWnd) <> 0 Then
    Select Case wMsg
        Case WM_MOUSEHOVER, WM_LBUTTONDOWN
            SetProp hWnd, StrPtr("Hot"), 1
            Button.Refresh
        Case WM_MOUSELEAVE
            SetProp hWnd, StrPtr("Hot"), 0
            Button.Refresh
        Case WM_KILLFOCUS
            Dim P As POINTAPI
            GetCursorPos P
            If WindowFromPoint(P.X, P.Y) <> hWnd Then SetProp hWnd, StrPtr("Hot"), 0
            Button.Refresh
        Case WM_MOUSEMOVE
            If GetProp(hWnd, StrPtr("Hot")) = 0 Then
                Dim TME As TRACKMOUSEEVENTSTRUCT
                With TME
                .cbSize = LenB(TME)
                .hWndTrack = hWnd
                .dwFlags = TME_HOVER Or TME_LEAVE
                .dwHoverTime = 1
                End With
                TrackMouseEvent TME
            End If
            If GetProp(hWnd, StrPtr("Painted")) = 0 Then Button.Refresh
        Case WM_SETFOCUS, WM_ENABLE
            If SetRedraw = True Then
                SendMessage hWnd, WM_SETREDRAW, 1, ByVal 0&
                If wMsg = WM_ENABLE Then
                    SetProp hWnd, StrPtr("Enabled"), 0
                    InvalidateRect hWnd, ByVal 0&, 0
                Else
                    SetProp hWnd, StrPtr("Enabled"), 1
                    Button.Refresh
                End If
            End If
        Case WM_LBUTTONUP, WM_RBUTTONUP
            Button.Refresh
    End Select
End If
End Function

Private Sub RemoveRedirectButton(ByVal hWnd As Long, ByVal uIdSubclass As Long)
RemoveWindowSubclass hWnd, AddressOf RedirectButton, uIdSubclass
End Sub

Private Sub DrawButton(ByVal hWnd As Long, ByVal hDC As Long, ByVal Button As Object)
Dim Theme As Long
Dim ButtonState As UxThemeButtonStates
Dim Enabled As Boolean, Checked As Boolean, Default As Boolean, Hot As Boolean, Focused As Boolean, Pushed As Boolean
Dim hFontOld As Long
Dim ButtonPicture As IPictureDisp, DisabledPictureAvailable As Boolean, ButtonFont As IFont
Dim ClientRect As RECT, TextRect As RECT
Dim RgnClip As Long
Dim CX As Long, CY As Long, X As Long, Y As Long
ButtonState = SendMessage(hWnd, BM_GETSTATE, 0, ByVal 0&)
Enabled = IIf(GetProp(hWnd, StrPtr("Enabled")) = 1, True, Button.Enabled)
Select Case TypeName(Button)
    Case "CommandButton", "CommandButtonW"
        Default = Button.Default
        If GetFocus() <> hWnd Then
            On Error Resume Next
            If CLng(Button.Parent.ActiveControl.Default) > 0 Then Else Default = False
            On Error GoTo 0
        End If
    Case "CheckBox", "CheckBoxW"
        Checked = IIf(Button.Value = vbChecked, True, False)
        Default = False
    Case "OptionButton", "OptionButtonW"
        Checked = Button.Value
        Default = False
End Select
Hot = IIf(GetProp(hWnd, StrPtr("Hot")) = 0, False, True)
If Checked = True Then Hot = False
Pushed = IIf((ButtonState And BST_PUSHED) = 0, False, True)
Focused = IIf((ButtonState And BST_FOCUS) = 0, False, True)
If Enabled = False Then
    ButtonState = PBS_DISABLED
    Set ButtonPicture = CoalescePicture(Button.DisabledPicture, Button.Picture)
    If Not Button.DisabledPicture Is Nothing Then
        If Button.DisabledPicture.Handle <> 0 Then DisabledPictureAvailable = True
    End If
ElseIf Hot = True And Pushed = False Then
    ButtonState = PBS_HOT
    If Checked = True Then
        Set ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
    Else
        Set ButtonPicture = Button.Picture
    End If
ElseIf Checked = True Or Pushed = True Then
    ButtonState = PBS_PRESSED
    Set ButtonPicture = CoalescePicture(Button.DownPicture, Button.Picture)
ElseIf Focused = True Or Default = True Then
    ButtonState = PBS_DEFAULTED
    Set ButtonPicture = Button.Picture
Else
    ButtonState = PBS_NORMAL
    Set ButtonPicture = Button.Picture
End If
If Not ButtonPicture Is Nothing Then
    If ButtonPicture.Handle = 0 Then Set ButtonPicture = Nothing
End If
GetClientRect hWnd, ClientRect
Theme = OpenThemeData(hWnd, StrPtr("Button"))
GetThemeBackgroundRegion Theme, hDC, BP_PUSHBUTTON, ButtonState, ClientRect, RgnClip
ExtSelectClipRgn hDC, RgnClip, RGN_DIFF
Call DrawRect(hDC, 0, 0, ClientRect.Right, ClientRect.Bottom, Button.BackColor)
If IsThemeBackgroundPartiallyTransparent(Theme, BP_PUSHBUTTON, ButtonState) <> 0 Then DrawThemeParentBackground hWnd, hDC, ClientRect
ExtSelectClipRgn hDC, 0, RGN_COPY
DeleteObject RgnClip
DrawThemeBackground Theme, hDC, BP_PUSHBUTTON, ButtonState, ClientRect, ClientRect
GetThemeBackgroundContentRect Theme, hDC, BP_PUSHBUTTON, ButtonState, ClientRect, ClientRect
If Focused = True Then DrawFocusRect hDC, ClientRect
If Not Button.Caption = vbNullString Or Len(Button.Caption) > 0 Then
    Set ButtonFont = Button.Font
    hFontOld = SelectObject(hDC, ButtonFont.hFont)
    LSet TextRect = ClientRect
    DrawText hDC, StrPtr(Button.Caption), -1, TextRect, DT_CALCRECT Or DT_WORDBREAK
    TextRect.Left = ClientRect.Left
    TextRect.Right = ClientRect.Right
    If ButtonPicture Is Nothing Then
        TextRect.Top = ((ClientRect.Bottom - TextRect.Bottom) / 2) + 3
        TextRect.Bottom = TextRect.Top + TextRect.Bottom
    Else
        TextRect.Top = (ClientRect.Bottom - TextRect.Bottom) + 1
        TextRect.Bottom = ClientRect.Bottom
    End If
    DrawThemeText Theme, hDC, BP_PUSHBUTTON, ButtonState, StrPtr(Button.Caption), -1, DT_CENTER Or DT_WORDBREAK, 0, TextRect
    SelectObject hDC, hFontOld
    ClientRect.Bottom = TextRect.Top
    ClientRect.Left = TextRect.Left
End If
CloseThemeData Theme
If Not ButtonPicture Is Nothing Then
    CX = CHimetricToPixel_X(ButtonPicture.Width)
    CY = CHimetricToPixel_Y(ButtonPicture.Height)
    X = ClientRect.Left + ((ClientRect.Right - ClientRect.Left - CX) / 2)
    Y = ClientRect.Top + ((ClientRect.Bottom - ClientRect.Top - CY) / 2)
    If Enabled = True Or DisabledPictureAvailable = True Then
        If ButtonPicture.Type = vbPicTypeBitmap And Button.UseMaskColor = True Then
            Dim hDCScreen As Long
            Dim hDC1 As Long, hBmpOld1 As Long
            hDCScreen = GetDC(0)
            If hDCScreen <> 0 Then
                hDC1 = CreateCompatibleDC(hDCScreen)
                If hDC1 <> 0 Then
                    hBmpOld1 = SelectObject(hDC1, ButtonPicture.Handle)
                    TransparentBlt hDC, X, Y, CX, CY, hDC1, 0, 0, CX, CY, WinColor(Button.MaskColor)
                    SelectObject hDC1, hBmpOld1
                    DeleteDC hDC1
                End If
                ReleaseDC 0, hDCScreen
            End If
        Else
            With ButtonPicture
            .Render hDC Or 0&, X Or 0&, Y + CY Or 0&, CX Or 0&, -CY Or 0&, 0&, 0&, .Width, .Height, ByVal 0&
            End With
        End If
    Else
        If ButtonPicture.Type = vbPicTypeIcon Then
            DrawState hDC, 0, 0, ButtonPicture.Handle, 0, X, Y, CX, CY, DST_ICON Or DSS_DISABLED
        ElseIf ButtonPicture.Type = vbPicTypeBitmap Then
            DrawState hDC, 0, 0, ButtonPicture.Handle, 0, X, Y, CX, CY, DST_BITMAP Or DSS_DISABLED
        End If
    End If
End If
End Sub

Private Sub DrawRect(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal FillColor As OLE_COLOR)
Dim RC As RECT
Dim Brush As Long
Brush = CreateSolidBrush(WinColor(FillColor))
With RC
.Left = X
.Top = Y
.Right = X + CX
.Bottom = Y + CY
End With
FillRect hDC, RC, Brush
DeleteObject Brush
End Sub

Private Function CoalescePicture(ByVal Picture As IPictureDisp, ByVal DefaultPicture As IPictureDisp) As IPictureDisp
If Picture Is Nothing Then
    Set CoalescePicture = DefaultPicture
ElseIf Picture.Handle = 0 Then
    Set CoalescePicture = DefaultPicture
Else
    Set CoalescePicture = Picture
End If
End Function
