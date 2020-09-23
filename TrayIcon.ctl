VERSION 5.00
Begin VB.UserControl TrayIcon 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   ClipControls    =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   420
   ToolboxBitmap   =   "TrayIcon.ctx":0000
   Begin VB.Image imgTray 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   525
      Picture         =   "TrayIcon.ctx":00FA
      Top             =   -30
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   420
      Left            =   0
      Picture         =   "TrayIcon.ctx":0244
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "TrayIcon Control"
Option Explicit

' Const needed to intercept mouse
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_MBUTTONDBLCLK = &H209

' Const needed as ID for the CallBackMessage
Private Const WM_MOUSEMOVE = &H200

' Needed for PostMessage
Private Const WM_NULL = &H0

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

' The TrayIcon structure
Private Type NOTIFYICONDATA
   cbSize As Long
   hWnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private tTray As NOTIFYICONDATA
Private bCreated As Boolean

Event MouseUp(Button As Integer)
Event MouseDown(Button As Integer)
Event MouseDblClk(Button As Integer)
Event ShowParent()                   ' When WM_LBUTTONDBLCLK
Event ShowPopUpMenu()                ' when WM_RBUTTONUP

'Public Sub ForceTaskSwitch()
'    If Not bCreated Then Exit Sub
'    ' Necessary to force task switch -- see Q135788
'    Call PostMessage(oOwner.hWnd, WM_NULL, 0, 0)
'End Sub

Private Sub UserControl_InitProperties()
    bCreated = False
    tTray.hIcon = imgTray.Picture
    tTray.szTip = vbNullChar
End Sub

'Force design-time control to size of icon
Private Sub UserControl_Resize()
    Size imgIcon.Width, imgIcon.Height
End Sub

Private Sub UserControl_Terminate()
    Destroy
End Sub

' TrayIcon only runs in Win95 or above (under NewShell).
Public Property Get HasValidShell() As Boolean
Attribute HasValidShell.VB_Description = "Returns True if current operating system will accept TrayIcons"
Attribute HasValidShell.VB_ProcData.VB_Invoke_Property = ";Misc"
    Dim os As OSVERSIONINFO
    os.dwOSVersionInfoSize = Len(os)
    Call GetVersionEx(os)
    HasValidShell = (os.dwMajorVersion > 3)
End Property

Public Function Create() As Boolean
Attribute Create.VB_Description = "Creates the TrayIcon"
    If bCreated Then Destroy

    If Not HasValidShell Then
        Create = False
        Exit Function
    End If

    tTray.cbSize = Len(tTray)
    tTray.hWnd = UserControl.hWnd
    tTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    tTray.uID = vbNull
    tTray.uCallbackMessage = WM_MOUSEMOVE

    Call Shell_NotifyIcon(NIM_ADD, tTray)
    bCreated = True
    Create = True
End Function

Public Sub Destroy()
Attribute Destroy.VB_Description = "Destroys the TrayIcon"
    If Not bCreated Then Exit Sub           ' You can only kill once....
    Call Shell_NotifyIcon(NIM_DELETE, tTray)
    bCreated = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bCreated Then Exit Sub           ' Ignore when not created

    On Error Resume Next

    Static lngMsg As Long
    Static blnFlag As Boolean

    lngMsg = X / Screen.TwipsPerPixelX

    If blnFlag = False Then
        blnFlag = True
        Select Case lngMsg
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton)
        Case WM_LBUTTONDBLCLK
            RaiseEvent MouseDblClk(vbLeftButton)
            RaiseEvent ShowParent               ' To make life easier
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
            RaiseEvent ShowPopUpMenu            ' To make life easier
        Case WM_RBUTTONDBLCLK
            RaiseEvent MouseDblClk(vbRightButton)
        Case WM_MBUTTONDOWN
            RaiseEvent MouseDown(vbMiddleButton)
        Case WM_MBUTTONUP
            RaiseEvent MouseUp(vbMiddleButton)
        Case WM_MBUTTONDBLCLK
            RaiseEvent MouseDblClk(vbMiddleButton)
        End Select

        blnFlag = False
    End If
End Sub

Public Property Get Active() As Boolean
Attribute Active.VB_Description = "Returns True when TrayIcon is created"
Attribute Active.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Active.VB_UserMemId = 0
    Active = bCreated
End Property

Public Property Get Icon() As Long
Attribute Icon.VB_Description = "Returns/Sets the icon displayed in the TrayIcon as run time"
Attribute Icon.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Icon.VB_MemberFlags = "400"
    Icon = tTray.hIcon
End Property

Public Property Let Icon(ByRef vNewValue As Long)
    tTray.hIcon = vNewValue

    If bCreated Then
        tTray.uFlags = NIF_ICON
        Call Shell_NotifyIcon(NIM_MODIFY, tTray)
    End If
End Property

Public Property Get Picture() As StdPicture
Attribute Picture.VB_Description = "Returns/Sets the picture displayed in the TrayIcon as run time"
Attribute Picture.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Picture.VB_MemberFlags = "200"
    Set Picture = imgTray.Picture
End Property

Public Property Let Picture(ByRef vNewValue As StdPicture)
    Set imgTray.Picture = vNewValue
    tTray.hIcon = imgTray.Picture

    If bCreated Then
        tTray.uFlags = NIF_ICON
        Call Shell_NotifyIcon(NIM_MODIFY, tTray)
    End If

    PropertyChanged "Picture"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/Sets the text displayed when the mouse is paused over the TrayIcon"
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Dim n As Integer
    n = InStr(tTray.szTip, vbNullChar)
    If n > 0 Then
        ToolTipText = Left(tTray.szTip, n - 1)
    Else
        ToolTipText = tTray.szTip
    End If
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
    tTray.szTip = vNewValue & vbNullChar

    If bCreated Then
        tTray.uFlags = NIF_TIP
        Call Shell_NotifyIcon(NIM_MODIFY, tTray)
    End If

    PropertyChanged "ToolTipText"
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Set imgTray.Picture = .ReadProperty("Picture", Nothing)
        tTray.szTip = .ReadProperty("ToolTipText", vbNullChar)
    End With

    If Not (imgTray.Picture Is Nothing) Then
        tTray.hIcon = imgTray.Picture
    End If
    If Right$(tTray.szTip, 1) <> vbNullChar Then
        tTray.szTip = tTray.szTip & vbNullChar
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Picture", imgTray.Picture, vbNull
        .WriteProperty "ToolTipText", tTray.szTip, vbNullChar
    End With
End Sub
