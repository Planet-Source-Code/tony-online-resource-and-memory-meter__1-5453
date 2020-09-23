VERSION 5.00
Begin VB.UserControl PopDown 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4500
   ControlContainer=   -1  'True
   ScaleHeight     =   1005
   ScaleWidth      =   4500
   ToolboxBitmap   =   "PopDown.ctx":0000
   Begin VB.CheckBox chkExpand 
      DownPicture     =   "PopDown.ctx":00FA
      Height          =   210
      Left            =   0
      MouseIcon       =   "PopDown.ctx":0404
      MousePointer    =   99  'Custom
      Picture         =   "PopDown.ctx":0556
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Value           =   1  'Checked
      Width           =   4500
   End
   Begin VB.PictureBox picFocus 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   165
      Left            =   30
      ScaleHeight     =   165
      ScaleWidth      =   180
      TabIndex        =   1
      Top             =   15
      Width           =   180
   End
End
Attribute VB_Name = "PopDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'#X-

' brainstorm: introduce Buddy-Controls.
' A Buddy-Control will automatic change the top value, when PopDown changes
' See updown control.

Public Enum PDBoxStyles
    None = 0
    Inset = 1
    Raised = 2
End Enum

Private Type PDSettingsType
    BoxStyle As PDBoxStyles
    BevelWidth As Integer
    SizeParent As Boolean
    Value As Boolean
    Height As Long
End Type

Private PDDev As PDSettingsType
Private mbRedraw As Boolean         ' Flag to disable redraw - usefull when updating lots of properties (speeds up when switching off, when done, switch back on again)
Private mbInitialised As Boolean

Event Change()

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = 3
    UserControl.AutoRedraw = True
    mbRedraw = False
End Sub

Private Sub UserControl_InitProperties()
    With PDDev
        .BoxStyle = None
        .BevelWidth = 0
        .SizeParent = True
        .Value = (chkExpand = vbChecked)
        .Height = UserControl.Height
    End With
    UserControl.BackColor = vbButtonFace
    picFocus.BackColor = UserControl.BackColor
    mbInitialised = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        PDDev.BoxStyle = .ReadProperty("BoxStyle", None)
        PDDev.BevelWidth = .ReadProperty("BevelWidth", 0)
        PDDev.SizeParent = .ReadProperty("SizeParent", True)
        UserControl.BackColor = .ReadProperty("BackColor", vbButtonFace)
    End With
    picFocus.BackColor = UserControl.BackColor
    PDDev.Value = (chkExpand = vbChecked)
    PDDev.Height = UserControl.Height
    mbInitialised = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BoxStyle", PDDev.BoxStyle, None
        .WriteProperty "BevelWidth", PDDev.BevelWidth, 0
        .WriteProperty "SizeParent", PDDev.SizeParent, True
        .WriteProperty "BackColor", UserControl.BackColor, vbButtonFace
    End With
    picFocus.BackColor = UserControl.BackColor
End Sub

Private Sub UserControl_Resize()
    Static bDejavu As Boolean
    If bDejavu Then Exit Sub
    bDejavu = True

    If UserControl.ScaleHeight < chkExpand.Height Then UserControl.Height = (chkExpand.Height * Screen.TwipsPerPixelY)
    chkExpand.Width = UserControl.ScaleWidth

    If chkExpand = vbChecked Then PDDev.Height = UserControl.Height

    UserControl_Paint
    bDejavu = False
End Sub

Private Sub UserControl_Paint()
    If PDDev.Value Then
        PaintBevel 0, chkExpand.Height + 1, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1, PDDev.BoxStyle, PDDev.BevelWidth
        If PDDev.BevelWidth < 1 Or PDDev.BoxStyle = None Then Exit Sub
        UserControl.Line (0, chkExpand.Height)-(UserControl.ScaleWidth - 1, chkExpand.Height), UserControl.Parent.BackColor
    End If
End Sub

Private Sub UserControl_Show()
    mbRedraw = True
    UserControl_Paint
End Sub

Private Sub chkExpand_Click()
    On Error Resume Next
    'UserControl.SetFocus
    picFocus.SetFocus

    Dim nHeight As Single
    nHeight = UserControl.Height    ' Previous height

    PDDev.Value = (chkExpand = vbChecked)

    If UserControl.Parent.WindowState = vbMaximized Then
        ' When the parent is maximized, other rules count
        If PDDev.Value Then     ' Expanded view (make bigger)
            UserControl.Height = PDDev.Height
        Else                    ' Collasped view (make smaller)
            UserControl.Height = (chkExpand.Height * Screen.TwipsPerPixelY)
        End If

        RaiseEvent Change               ' Tell'em afterwards - so parent can adjust other child controls.
    
    Else
        RaiseEvent Change               ' Tell'em before it starts - so parent can prepare for it.
    
        If PDDev.Value Then     ' Expanded view (make bigger)
            UserControl.Height = PDDev.Height
            If PDDev.SizeParent Then UserControl.Parent.Height = UserControl.Parent.Height + (UserControl.Height - nHeight)
        Else                    ' Collasped view (make smaller)
            UserControl.Height = (chkExpand.Height * Screen.TwipsPerPixelY)
            If PDDev.SizeParent Then UserControl.Parent.Height = UserControl.Parent.Height - (nHeight - UserControl.Height)
        End If
    End If
End Sub

' ---------------------------------------------------------------
' Internal generic painting routines

Private Sub PaintBevel(ByVal nStartX As Integer, ByVal nStartY As Integer, ByVal nEndX As Integer, ByVal nEndY As Integer, ByVal nStyle As PDBoxStyles, ByVal nBevelWidth As Integer)
    Dim i As Integer, nBevel As Integer

    If nBevelWidth < 1 Or nStyle = None Then Exit Sub

    With UserControl
        .FillStyle = 0
        .DrawStyle = 0
        .DrawWidth = 1
    End With

    UserControl.Line (nStartX, nStartY)-(nEndX, nEndY), UserControl.BackColor, BF

    nBevel = nBevelWidth - 1

    Select Case nStyle  ' Paint bevel
    Case Raised
        ' Enhance the outside
        UserControl.Line (nStartX, nStartY)-(nEndX - 1, nStartY), vb3DLight
        UserControl.Line (nStartX, nStartY)-(nStartX, nEndY - 1), vb3DLight
        UserControl.Line (nStartX, nEndY)-(nEndX + 1, nEndY), vb3DDKShadow ' vbBlack
        UserControl.Line (nEndX, nStartY)-(nEndX, nEndY), vb3DDKShadow ' vbBlack

        ' Paint the shadow
        For i = 1 To nBevel
            UserControl.Line (nStartX + i, nStartY + i)-(nEndX - i, nStartY + i), vb3DHighlight ' RGB(255, 255, 255)
            UserControl.Line (nStartX + i, nStartY + i)-(nStartX + i, nEndY - i), vb3DHighlight ' RGB(255, 255, 255)
            UserControl.Line (nStartX + i, nEndY - i)-(nEndX - i + 1, nEndY - i), vbButtonShadow ' RGB(92, 92, 92)
            UserControl.Line (nEndX - i, nStartY + i)-(nEndX - i, nEndY - i), vbButtonShadow ' RGB(92, 92, 92)
        Next i

    Case Inset
        ' Paint the shadow
        For i = 0 To (nBevel - 1)
            UserControl.Line (nStartX + i, nStartY + i)-(nEndX - i, nStartY + i), vbButtonShadow
            UserControl.Line (nStartX + i, nStartY + i)-(nStartX + i, nEndY - i), vbButtonShadow
            UserControl.Line (nStartX + i, nEndY - i)-(nEndX - i + 1, nEndY - i), vb3DHighlight
            UserControl.Line (nEndX - i, nStartY + i)-(nEndX - i, nEndY - i + 1), vb3DHighlight
        Next i

        ' Enhance the inside
        UserControl.Line (nStartX + nBevel, nStartY + nBevel)-(nEndX - nBevel, nStartY + nBevel), vb3DDKShadow
        UserControl.Line (nStartX + nBevel, nStartY + nBevel)-(nStartX + nBevel, nEndY - nBevel), vb3DDKShadow
        UserControl.Line (nStartX + nBevel, nEndY - nBevel)-(nEndX - nBevel + 1, nEndY - nBevel), vb3DLight
        UserControl.Line (nEndX - nBevel, nStartY + nBevel)-(nEndX - nBevel, nEndY - nBevel + 1), vb3DLight
    End Select
End Sub

' ---------------------------------------------------------------

Public Function Expanded() As Boolean
    Expanded = PDDev.Value
End Function

' ---------------------------------------------------------------
' Properties

Public Property Get Value() As Boolean
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0
    Value = PDDev.Value
End Property
Public Property Let Value(ByVal vNewValue As Boolean)
    chkExpand = IIf(vNewValue, vbChecked, vbUnchecked)
End Property

Public Property Get BoxStyle() As PDBoxStyles
Attribute BoxStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BoxStyle = PDDev.BoxStyle
End Property
Public Property Let BoxStyle(ByVal vNewValue As PDBoxStyles)
    PDDev.BoxStyle = vNewValue
    UserControl_Paint
    PropertyChanged "BoxStyle"
End Property

Public Property Get BevelWidth() As Integer
Attribute BevelWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BevelWidth = PDDev.BevelWidth
End Property
Public Property Let BevelWidth(ByVal vNewValue As Integer)
    If vNewValue < 0 Then Exit Property
    PDDev.BevelWidth = vNewValue
    UserControl_Paint
    PropertyChanged "BevelWidth"
End Property

Public Property Get BackColor() As Long
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As Long)
    UserControl.BackColor = vNewValue
    picFocus.BackColor = UserControl.BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get CanvasHeight() As Single
Attribute CanvasHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CanvasHeight = PDDev.Height - (chkExpand.Height * Screen.TwipsPerPixelY)
End Property

Public Property Let CanvasHeight(ByVal vNewValue As Single)
    If vNewValue < 0 Then Exit Property
    PDDev.Height = vNewValue + (chkExpand.Height * Screen.TwipsPerPixelY)
    chkExpand_Click
End Property

Public Property Get ButtonHeight() As Single
Attribute ButtonHeight.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ButtonHeight = chkExpand.Height * Screen.TwipsPerPixelY
End Property

Public Property Let ButtonHeight(ByVal vNewValue As Single)
    If vNewValue < 0 Then Exit Property
    Dim nHeight As Single
    nHeight = chkExpand.Height
    chkExpand.Height = vNewValue / Screen.TwipsPerPixelY
    If PDDev.Value Then
        If PDDev.Height < (chkExpand.Height * Screen.TwipsPerPixelY) Then PDDev.Height = chkExpand.Height * Screen.TwipsPerPixelY
    Else
        If UserControl.ScaleHeight < chkExpand.Height Then UserControl.Height = (chkExpand.Height * Screen.TwipsPerPixelY)
        PDDev.Height = PDDev.Height + ((nHeight - chkExpand.Height) * Screen.TwipsPerPixelY)
        If PDDev.SizeParent Then UserControl.Parent.Height = UserControl.Parent.Height - ((nHeight - chkExpand.Height) * Screen.TwipsPerPixelY)
    End If
    If PDDev.Height < 0 Then PDDev.Height = 0
    chkExpand_Click
End Property

Public Property Get AutoSizeParent() As Boolean
Attribute AutoSizeParent.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AutoSizeParent = PDDev.SizeParent
End Property

Public Property Let AutoSizeParent(ByVal vNewValue As Boolean)
    PDDev.SizeParent = vNewValue
End Property
