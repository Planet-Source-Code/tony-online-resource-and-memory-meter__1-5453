VERSION 5.00
Begin VB.Form frmResource 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3510
   ControlBox      =   0   'False
   Icon            =   "Resource.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ResourceMeter.CoolButton btnMark 
      Height          =   210
      Left            =   3300
      TabIndex        =   16
      ToolTipText     =   "Mark current percentage"
      Top             =   210
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Resource.frx":000C
      PictureDown     =   "Resource.frx":02C6
      PictureOver     =   "Resource.frx":0580
      colourTextOverColour=   8388608
      UseBevels       =   0   'False
      UseUnderlineOnFocus=   0   'False
      TextPositionV   =   0
      TextPositionH   =   0
   End
   Begin ResourceMeter.PopDown PopDown 
      Height          =   1845
      Left            =   3300
      TabIndex        =   1
      ToolTipText     =   "Show/Hide memory information"
      Top             =   420
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   3254
   End
   Begin ResourceMeter.CoolButton btnClose 
      Height          =   210
      Left            =   3300
      TabIndex        =   0
      ToolTipText     =   "Close view"
      Top             =   0
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Resource.frx":083A
      PictureDown     =   "Resource.frx":0AF4
      PictureOver     =   "Resource.frx":0DAE
      colourTextOverColour=   8388608
      UseBevels       =   0   'False
      UseUnderlineOnFocus=   0   'False
      TextPositionV   =   0
      TextPositionH   =   0
   End
   Begin ResourceMeter.TrayIcon TrayIcon 
      Left            =   525
      Top             =   2610
      _ExtentX        =   741
      _ExtentY        =   741
      Picture         =   "Resource.frx":1068
      ToolTipText     =   "                                                                "
   End
   Begin ResourceMeter.ProgressGuage guageSystem 
      Height          =   210
      Left            =   0
      ToolTipText     =   "Percentage free system resources"
      Top             =   0
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   370
      Alignment       =   0
      Suffix          =   "% System"
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   30
      Top             =   2595
   End
   Begin ResourceMeter.ProgressGuage guageUser 
      Height          =   210
      Left            =   0
      ToolTipText     =   "Percentage free user resources"
      Top             =   210
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   370
      Alignment       =   0
      Suffix          =   "% User"
   End
   Begin ResourceMeter.ProgressGuage guageGDI 
      Height          =   210
      Left            =   0
      ToolTipText     =   "Percentage free GDI resources"
      Top             =   420
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   370
      Alignment       =   0
      Suffix          =   "% GDI"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Physical Memory:"
      Height          =   225
      Left            =   330
      TabIndex        =   15
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Available Physical Memory:"
      Height          =   225
      Left            =   90
      TabIndex        =   14
      Top             =   945
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Virtual Memory:"
      Height          =   225
      Left            =   210
      TabIndex        =   13
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Page File:"
      Height          =   225
      Left            =   690
      TabIndex        =   12
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Available Page File:"
      Height          =   225
      Left            =   90
      TabIndex        =   11
      Top             =   1845
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Memory Load:"
      Height          =   225
      Left            =   690
      TabIndex        =   10
      Top             =   2055
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2130
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2130
      TabIndex        =   8
      Top             =   945
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2130
      TabIndex        =   7
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   2130
      TabIndex        =   6
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   2130
      TabIndex        =   5
      Top             =   1845
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   2130
      TabIndex        =   4
      Top             =   2055
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Available Virtual Memory:"
      Height          =   225
      Left            =   210
      TabIndex        =   3
      Top             =   1395
      Width           =   1815
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "n/a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   2130
      TabIndex        =   2
      Top             =   1395
      Width           =   1095
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   12
      Left            =   1095
      Picture         =   "Resource.frx":11C2
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   11
      Left            =   765
      Picture         =   "Resource.frx":130C
      Stretch         =   -1  'True
      Top             =   3495
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   10
      Left            =   420
      Picture         =   "Resource.frx":1456
      Stretch         =   -1  'True
      Top             =   3495
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   9
      Left            =   90
      Picture         =   "Resource.frx":15A0
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   8
      Left            =   2640
      Picture         =   "Resource.frx":16EA
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   7
      Left            =   2280
      Picture         =   "Resource.frx":1834
      Stretch         =   -1  'True
      Top             =   3105
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   6
      Left            =   1980
      Picture         =   "Resource.frx":197E
      Stretch         =   -1  'True
      Top             =   3090
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   5
      Left            =   1665
      Picture         =   "Resource.frx":1AC8
      Stretch         =   -1  'True
      Top             =   3105
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   4
      Left            =   1347
      Picture         =   "Resource.frx":1C12
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   3
      Left            =   1029
      Picture         =   "Resource.frx":1D5C
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   2
      Left            =   711
      Picture         =   "Resource.frx":1EA6
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   1
      Left            =   375
      Picture         =   "Resource.frx":1FF0
      Stretch         =   -1  'True
      Top             =   3135
      Width           =   240
   End
   Begin VB.Image imgMeter 
      Height          =   240
      Index           =   0
      Left            =   75
      Picture         =   "Resource.frx":213A
      Stretch         =   -1  'True
      Top             =   3135
      Width           =   240
   End
   Begin VB.Menu mnuPopUpMenu 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuDisplay 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SR = 0
Private Const GDI = 1
Private Const USR = 2

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Type MEMORYSTATUS
    dwLength        As Long ' sizeof(MEMORYSTATUS)
    dwMemoryLoad    As Long ' percent of memory in use
    dwTotalPhys     As Long ' bytes of physical memory
    dwAvailPhys     As Long ' free physical memory bytes
    dwTotalPageFile As Long ' bytes of paging file
    dwAvailPageFile As Long ' free bytes of paging file
    dwTotalVirtual  As Long ' user bytes of address space
    dwAvailVirtual  As Long ' free user bytes
End Type

'Loads a MEMORYSTATUS structure with information about the current state of the systems memory.
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private mIsWin32 As Boolean

' Used in FormOnTop (in Form_Load and Form_Unload)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long

' Drag form
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Private Sub Form_Load()
    Dim OSInfo As OSVERSIONINFO

    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1

    PopDown.Value = False

    ' Operating System/Resource Information.
    '
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    Call GetVersionEx(OSInfo)
    mIsWin32 = (OSInfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS)

    tmrUpdate_Timer
    tmrUpdate.Enabled = True

    If Not mIsWin32 Then
        With TrayIcon
            .Icon = Me.Icon
            .ToolTipText = "Unable to display resources usage"
        End With
        PopDown.Value = True
    End If

    If TrayIcon.Create Then Me.Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then FormDrag
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If TrayIcon.Active And UnloadMode = vbFormControlMenu Then
      Me.Visible = False
      Cancel = True
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, 2 Or 1
    TrayIcon.Destroy
End Sub

Private Sub guageSystem_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then FormDrag
End Sub

Private Sub guageUser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then FormDrag
End Sub

Private Sub guageGDI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then FormDrag
End Sub

Private Sub PopDown_Change()
    tmrUpdate_Timer
End Sub

Private Sub tmrUpdate_Timer()
    Dim MS As MEMORYSTATUS

    On Local Error Resume Next

    If PopDown.Value Then
        MS.dwLength = Len(MS)
        Call GlobalMemoryStatus(MS)
        With MS
            lbl(0) = Format$(.dwTotalPhys / 1024, "#,###") & " Kb"
            lbl(1) = Format$(.dwAvailPhys / 1024, "#,###") & " Kb"
            lbl(2) = Format$(.dwTotalVirtual / 1024, "#,###") & " Kb"
            lbl(3) = Format$(.dwAvailVirtual / 1024, "#,###") & " Kb"
            lbl(4) = Format$(.dwTotalPageFile / 1024, "#,###") & " Kb"
            lbl(5) = Format$(.dwAvailPageFile / 1024, "#,###") & " Kb"
            lbl(6) = Format$(.dwMemoryLoad, "##0") & "%"
        End With
    End If

    If mIsWin32 Then
        Dim nIndex As Integer

        guageSystem.Value = pBGetFreeSystemResources(SR)
        guageSystem.ForeColor = AssignColour(guageSystem.Value)
        guageUser.Value = pBGetFreeSystemResources(USR)
        guageUser.ForeColor = AssignColour(guageUser.Value)
        guageGDI.Value = pBGetFreeSystemResources(GDI)
        guageGDI.ForeColor = AssignColour(guageGDI.Value)

        Select Case guageSystem.Value
        Case Is > 99
            TrayIcon.Icon = imgMeter(0).Picture
        Case Is < 1
            TrayIcon.Icon = imgMeter(12).Picture
        Case Else
            nIndex = (100 - guageSystem.Value) * 0.13
            Select Case nIndex
            Case Is < 0             ' Empty
                nIndex = 0
            Case Is > 12            ' Full
                nIndex = 12
            End Select
            TrayIcon.Icon = imgMeter(nIndex).Picture
        End Select
        TrayIcon.ToolTipText = "System: " & guageSystem.Value & "% " & _
                               "User: " & guageUser.Value & "% " & _
                               "GDI: " & guageGDI.Value & "%"
    End If
End Sub

Private Function AssignColour(nPercent As Integer) As Long
    Select Case nPercent
    Case Is < 10
        AssignColour = vbRed
    Case Is < 25
        AssignColour = vbMagenta
    Case Else
        AssignColour = vbBlue
    End Select
End Function

Private Sub FormDrag()
    ReleaseCapture
    Call SendMessage(Me.hwnd, &HA1, 2, 0&)
End Sub

Private Sub btnClose_Click()
    Me.Visible = False
End Sub

Private Sub btnMark_Click()
    With guageSystem
        .MarkValue = .Value
        .UseMark = True
        .ToolTipText = "Percentage free system resources (Marked at " & .MarkValue & "%)"
    End With
    With guageUser
        .MarkValue = .Value
        .UseMark = True
        .ToolTipText = "Percentage free user resources (Marked at " & .MarkValue & "%)"
    End With
    With guageGDI
        .MarkValue = .Value
        .UseMark = True
        .ToolTipText = "Percentage free GDI resources (Marked at " & .MarkValue & "%)"
    End With
End Sub

Private Sub mnuDisplay_Click()
    Me.Visible = Not Me.Visible
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub TrayIcon_ShowParent()
    Me.Visible = True
End Sub

Private Sub TrayIcon_ShowPopUpMenu()
    mnuDisplay.Checked = Me.Visible
    PopupMenu mnuPopUpMenu, , , , mnuDisplay
End Sub
