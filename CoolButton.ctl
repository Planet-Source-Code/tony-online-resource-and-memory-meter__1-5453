VERSION 5.00
Begin VB.UserControl CoolButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
   EditAtDesignTime=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   28
   ToolboxBitmap   =   "CoolButton.ctx":0000
End
Attribute VB_Name = "CoolButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------
'                CoolButton control, ver 2.2
'
' (C) Dave Hng '99                   starfox@earthcorp.com
'----------------------------------------------------------
'
'A lot nicer with regards to system resources and CPU time,
'using SetCapture and ReleaseCapture instead of a timer,
'though a shit lot more confusing, especially the DrawBevel sub. :)
'
'Files for this usercontrol:
'----------------------------------------------------------
'cSfCoolButton.ctl
'
'Nothing else! Add it, and off you go!
'
'Known problems:
'----------------------------------------------------------
'Tooltips don't agree with SetCapture, it doesn't display them.
' -Can be rectified through subclassing, but that's a lot of work.
'Bevels are not drawn when in design mode, because i don't want to change lots of subs and functions.
' -it works, i'm not going to break it again.. :)
'Never name a property TextFont, it won't work for some reason.. :P
' -Causes problems, property is never saved.. odd.
'AutoDim doesn't work all the time
' -Don't know why.

'----------------------------------------------------------
'You shouldn't need to modify anything below here...
'(You shouldn't need to modify anything at all.. :) )

Option Explicit

'Constants for AutoDim.
Private Const csDimPercent As Single = 0.9 'Dim to 90%
Private Const csBriPercent As Single = 1.2 'Brighten to 120%
Private Const cbMaxValue As Byte = 255     'Max value for a byte

'API Declares
'I'm sure this is easier with VC++ and MFC..., though this is probably FASTER. :)
'----------------------------------------------------------
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CreateHalftonePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hdc As Long, ByVal wFlags As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyHeight As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

'You might like to use this function instead of CreateCompatibleBitmap, if it doesn't work for some reason.
'Private Declare Function CreateDiscardableBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Const DI_NORMAL = &H3       ' Needed for DrawIconEx

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'Constants for API calls
'----------------------------------------------------------
Private Const TA_CENTER = 6
Private Const TA_LEFT = 0
Private Const TA_RIGHT = 2
Private Const TA_BASELINE = 24

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const TRANSPARENT = 1

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0& '  color table in RGBs

'TypeDef Structs that this control uses
'----------------------------------------------------------
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Enum eBevelType
    '----------------------------------------------------------
    'Do not change these values, they are set for specific reasons,
    'as i do some bit operations on them to change settings.
    'It works like this, each value is two bits:
    '
    '           1                      1
    '     Mouse Up or Down       Mouse in area?
    '      -0 if Up, 1 if Down    -0 if Out, 1 if In
    '
    'Heh, and you thought VB programmers never knew what bits were.. :)
    '----------------------------------------------------------
    UpIn = 1
    DownIn = 3
    UpOut = 0
    DownOut = 2
End Enum

'Bevel width constant
Private Const ciBevelWidth  As Integer = 1

Public Enum eVTextPosition
    cTop = 0
    cMiddle = 1
    cBottom = 2
    c3Quarters = 3
End Enum

Public Enum eHTextPosition
    ciLeft = 0
    ciCenter = 1
    ciRight = 2
End Enum

'Property variables
Private bLoaded As Boolean
Private bUnderlineFocus As Boolean
Private bUsePictures As Boolean
Private bUseBevels As Boolean
Private bDipControls As Boolean
Private iBevelType As eBevelType
Private bDeviated As Boolean
Private iInitialScaleMode As Integer
Private bAutoSize As Boolean
Private sCaption As String
Private bEnabled As Boolean
Private bButtonsAlwaysUp As Boolean
Private bAutoDim As Boolean
Private lvTextPosition As Long
Private lhTextPosition As Long
Private bAutoColour As Boolean

Private hMouseOverBitmap As Long
Private hMouseDownBitmap As Long

'Pictures
Private picNormal As StdPicture
Private picMouseOver As StdPicture
Private picMouseDown As StdPicture

'Colours!
Private colour_Highlight As OLE_COLOR
Private colour_LowLight As OLE_COLOR
Private colour_BackColour As OLE_COLOR
Private colour_TextStdColour As OLE_COLOR
Private colour_TextOverColour As OLE_COLOR
Private colour_Ignore As OLE_COLOR

'Working variables
Private ti As Integer
Private ti2 As Integer
Private bClick As Boolean
Private bMouseDowned As Boolean

'Events
'----------------------------------------------------------
Public Event Click()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseExit()

Private Sub UserControl_Initialize()
    'Set initial values for variables that i can.
    '----------------------------------------------------------
    iBevelType = UpOut
    iInitialScaleMode = UserControl.ScaleMode

    colourHighlight = QBColor(15)
    colourLowLight = QBColor(8)
    colourBackColour = vbButtonFace
    colourTextStdColour = QBColor(0)
    colourTextOverColour = QBColor(1)

    UseBevels = True
    UsePictures = True
    bDipControls = False
    AutoSize = False
    UseUnderlineOnFocus = True
    bEnabled = True
End Sub

Private Sub UserControl_Terminate()
    FreeDimmedBitmaps
End Sub

Private Sub UserControl_Resize()
    DrawBevel iBevelType
End Sub

Private Sub UserControl_Show()
    DrawBevel iBevelType
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Dim picTemp As StdPicture

    With PropBag
        Set picNormal = .ReadProperty("Picture", picTemp)
        Set picMouseDown = .ReadProperty("PictureDown", picTemp)
        Set picMouseOver = .ReadProperty("PictureOver", picTemp)
        colourHighlight = .ReadProperty("colourHighlight", QBColor(15))
        colourLowLight = .ReadProperty("colourLowlight", QBColor(8))
        colourBackColour = .ReadProperty("colourBackColour", vbButtonFace)
        colourTextStdColour = .ReadProperty("colourTextStdColour", QBColor(0))
        colourTextOverColour = .ReadProperty("colourTextOverColour", colour_TextStdColour)
        colourIgnore = .ReadProperty("colourIgnore", vbBlack)
        Caption = .ReadProperty("Caption", "")
        UseBevels = .ReadProperty("UseBevels", True)
        UsePictures = .ReadProperty("UsePictures", True)
        UseDippedControls = .ReadProperty("UseDippedControls", False)
        AutoSize = .ReadProperty("AutoSize", False)
        UseUnderlineOnFocus = .ReadProperty("UseUnderlineOnFocus", True)
        Enabled = .ReadProperty("Enabled", True)
        Set UserControl.Font = .ReadProperty("CaptionFont", UserControl.Font)
        bButtonsAlwaysUp = .ReadProperty("AlwaysDrawBevel", False)
        AutoDim = .ReadProperty("AutoDim", False)
        TextPositionV = .ReadProperty("TextPositionV", cMiddle)
        TextPositionH = .ReadProperty("TextPositionH", ciCenter)
        AutoColour = .ReadProperty("AutoColour", False)
    End With

    UserControl.BackColor = colour_BackColour
    bLoaded = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Dim picTemp As StdPicture, fntTemp As Font

    With PropBag
        .WriteProperty "CaptionFont", UserControl.Font, fntTemp
        .WriteProperty "Picture", picNormal, picTemp
        .WriteProperty "PictureDown", picMouseDown, picTemp
        .WriteProperty "PictureOver", picMouseOver, picTemp
        .WriteProperty "colourHighlight", colour_Highlight, QBColor(15)
        .WriteProperty "colourLowlight", colour_LowLight, QBColor(8)
        .WriteProperty "colourBackColour", colour_BackColour, &H8000000F
        .WriteProperty "colourTextStdColour", colour_TextStdColour, QBColor(0)
        .WriteProperty "colourTextOverColour", colour_TextOverColour, colour_TextStdColour
        .WriteProperty "colourIgnore", colourIgnore, vbBlack
        .WriteProperty "Caption", sCaption, ""
        .WriteProperty "UseBevels", UseBevels, True
        .WriteProperty "UsePictures", UsePictures, True
        .WriteProperty "UseDippedControls", UseDippedControls, False
        .WriteProperty "AutoSize", AutoSize, False
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "UseUnderlineOnFocus", UseUnderlineOnFocus, True
        .WriteProperty "AlwaysDrawBevel", bButtonsAlwaysUp, False
        .WriteProperty "AutoDim", AutoDim, False
        .WriteProperty "TextPositionV", TextPositionV, cMiddle
        .WriteProperty "TextPositionH", TextPositionH, ciCenter
        .WriteProperty "AutoColour", AutoColour, False
    End With
End Sub

Private Sub UserControl_EnterFocus()
    UserControl.FontUnderline = bUnderlineFocus
End Sub

Private Sub UserControl_ExitFocus()
    UserControl.FontUnderline = False
    If bUnderlineFocus Then
        DrawBevel iBevelType
    End If
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not bEnabled Then Exit Sub
    'Traps for spacebar, if it's pushed, then behave like a button
    '----------------------------------------------------------
    If KeyCode = ti2 Then Exit Sub

    If KeyCode = vbKeySpace Then
        ti = iBevelType
        iBevelType = DownIn
        DrawBevel iBevelType
        UserControl.Refresh
    End If

    ti2 = KeyCode
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If Not bEnabled Then Exit Sub
    'If enter / return 's pressed, then simulate the button going
    'up, then down.
    '----------------------------------------------------------
    If KeyAscii = vbKeyReturn Then
        Dim iPrevBeveltype

        iPrevBeveltype = iBevelType

        iBevelType = DownIn
        DrawBevel iBevelType
        UserControl.Refresh

        Sleep 50

        iBevelType = UpIn
        DrawBevel iBevelType
        UserControl.Refresh

        Sleep 50

        RaiseEvent Click

        iBevelType = iPrevBeveltype
        DrawBevel iBevelType

    End If
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If Not bEnabled Then Exit Sub
    'Accompanying part for the KeyDown sub
    '----------------------------------------------------------
    If KeyCode = vbKeySpace And ti2 = vbKeySpace Then
        iBevelType = UpIn
        DrawBevel (iBevelType)
        UserControl.Refresh

        Sleep 50

        RaiseEvent Click

        iBevelType = ti
        ti = 0
        DrawBevel (iBevelType)
        ti2 = 0
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bEnabled Then Exit Sub

    Dim result As Long
    Dim bInArea As Boolean

    bInArea = ((X >= UserControl.ScaleLeft And X <= UserControl.ScaleWidth) And (Y >= UserControl.ScaleTop And Y <= UserControl.ScaleHeight))

    bClick = False

    If Button = vbLeftButton Then
        bMouseDowned = True
        'Mouse down, in area

        iBevelType = iBevelType Or 2
        DrawBevel iBevelType

        If (iBevelType = UpIn Or iBevelType = DownIn) Then
            result = SetCapture(UserControl.hwnd)
        End If

        bClick = (iBevelType And 1 = 1)

        bDeviated = True

    ElseIf Button = vbRightButton Then
        'Redraw with the mouse out.
        'iBevelType = UpOut
        'DrawBevel iBevelType
    End If

    If bInArea Then
        RaiseEvent MouseDown(Button, Shift, X, Y)
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bEnabled Then Exit Sub

    Dim result As Long
    Dim iPrevBevel As Integer
    Dim bInArea As Boolean

    'Bug / Glitch: VB doesn't update X and Y for a scalemode
    'If you change scalemode in the sub, X and Y are not changed, ever!

    UserControl.ScaleMode = iInitialScaleMode

    If Button = 0 Then
        iBevelType = iBevelType And 1
    ElseIf Button = vbLeftButton And bMouseDowned Then
        iBevelType = iBevelType Or 2
    End If

    iPrevBevel = iBevelType

    bInArea = ((X >= UserControl.ScaleLeft And X <= UserControl.ScaleWidth) And (Y >= UserControl.ScaleTop And Y <= UserControl.ScaleHeight))

    If bInArea Then
        'Set iBevelType to reflect that the mouse is in
        iBevelType = iBevelType Or 1
    Else
        'Set iBeveltype to reflect that the mouse is out
        iBevelType = iBevelType And 2
    End If

    If (iBevelType And 1) Then
        'Debug.Print "mouse in area"

        If iPrevBevel <> iBevelType Then
            DrawBevel iBevelType

            'MouseEnter is raised here, only occurs once.
            RaiseEvent MouseEnter
            result = SetCapture(UserControl.hwnd)
        End If

        RaiseEvent MouseMove(Button, Shift, X, Y)

    Else
        'I can raise the event here, because it'll only get called
        'once, before the usercontrol releases capture of mouse events.

        RaiseEvent MouseExit

        iBevelType = UpOut
        DrawBevel iBevelType
        result = ReleaseCapture()
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not bEnabled Then Exit Sub

    Dim result As Long
    Dim bInArea As Boolean

    bInArea = ((X >= UserControl.ScaleLeft And X <= UserControl.ScaleWidth) And (Y >= UserControl.ScaleTop And Y <= UserControl.ScaleHeight))

    'VB releases capture on mouseup somehow...,
    'might be how it's coded.

    If Button = vbRightButton Then
        result = SetCapture(UserControl.hwnd)
    End If

    If Button = vbLeftButton Then

        iBevelType = iBevelType And 1
        DrawBevel iBevelType

        result = SetCapture(UserControl.hwnd)

        bDeviated = False

    End If

    If bClick And (iBevelType And 1 = 1) And bMouseDowned Then
        bClick = False
        RaiseEvent Click
    End If

    If bInArea Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
    End If

    If Button = vbLeftButton Then result = SetCapture(UserControl.hwnd)
    bMouseDowned = False
End Sub

Private Sub AutoSizeControl()
    Dim result As Long, bmp As BITMAP

    'Find bitmap's dimensions. I don't know what picture
    'object width and height is measured in... something weird.
    result = GetObject(picNormal.Handle, Len(bmp), bmp)

    UserControl.ScaleMode = vbPixels

    'Leave room for bevels if needed
    If bUseBevels Then
        UserControl.Height = (bmp.bmHeight + 2) * Screen.TwipsPerPixelY
        UserControl.Width = (bmp.bmWidth + 2) * Screen.TwipsPerPixelX
    Else
        UserControl.Height = bmp.bmHeight * Screen.TwipsPerPixelY
        UserControl.Width = bmp.bmWidth * Screen.TwipsPerPixelX
    End If
End Sub

Private Sub DrawBevel(ByVal nBevelType As Integer)
    On Error GoTo ErrorHandler

    'Exit this sub if things aren't loaded, otherwise trouble will arise
    If Not bLoaded Then Exit Sub

    'Manual bitmap drawing, and text output!
    'Sheesh, what a waste of time :)
    'You can't use image and label controls, because they receive
    'mouse events, rather than the control, which messes things up.
    '------------------------------------------------------------------

    Dim result As Long, ts As String
    Dim picDraw As StdPicture
    Dim hBitmapHack As Long
    Dim bInnerBevel As Boolean
    Dim bBevel As Boolean

    UserControl.ScaleMode = vbPixels
    UserControl.Cls

    'Set vars appropriately
    '------------------------------------------------------------------
    Select Case nBevelType
    Case DownOut
        bInnerBevel = True
        bBevel = True
        If bUsePictures Then Set picDraw = picNormal
        UserControl.ForeColor = colour_TextStdColour
        hBitmapHack = picNormal.Handle

    Case DownIn
        bInnerBevel = True
        bBevel = True
        If bUsePictures Then Set picDraw = picMouseDown
        UserControl.ForeColor = colour_TextOverColour
        hBitmapHack = hMouseDownBitmap

    Case UpIn
DrawUp:
        UserControl.Cls
        bInnerBevel = False
        bBevel = bUseBevels
        If (bUsePictures And Not (picMouseOver Is Nothing)) Then Set picDraw = picMouseOver
        UserControl.ForeColor = colour_TextOverColour
        hBitmapHack = hMouseOverBitmap

    Case UpOut
        If bButtonsAlwaysUp Then GoTo DrawUp
        bBevel = False
        UserControl.Cls
        If bUsePictures Then Set picDraw = picNormal
        UserControl.ForeColor = colour_TextStdColour
        hBitmapHack = picNormal.Handle

    End Select

    'Check in case there's no picture, if not, bail.
    If picDraw Is Nothing Then Set picDraw = picNormal
    If picDraw.Handle = 0 Then Exit Sub

    'This next part draws the image and text to the usercontrol
    'I seriously hope there are no memory leaks here.
    '------------------------------------------------------------------
    Dim dcDesktop As Long, palHalfTone As Long
    Dim dcTemp As Long, palOld As Long
    Dim bmpOld As Long, bmp As BITMAP, rt As RECT
    Dim XPos As Long, YPos As Long
    Dim oldTextAlign As Long
    Dim oldTextDrawMode As Long

    'Create a halftone palette to dither to, if needed.
    palHalfTone = CreateHalftonePalette(UserControl.hdc)

    'Create off screen DC to draw to
    dcDesktop = GetDC(ByVal 0&)
    dcTemp = CreateCompatibleDC(dcDesktop)
    palOld = SelectPalette(dcTemp, palHalfTone, True)
    RealizePalette dcTemp

    'Associate picture with dc, including self generated dimmed bitmaps
    If bAutoDim Then
        bmpOld = SelectObject(dcTemp, hBitmapHack)
    Else
        bmpOld = SelectObject(dcTemp, picDraw.Handle)
    End If

    'Blit picture to usercontrol's center
    result = GetObject(picDraw.Handle, Len(bmp), bmp)
    XPos = UserControl.ScaleWidth / 2 - bmp.bmWidth / 2
    YPos = UserControl.ScaleHeight / 2 - bmp.bmHeight / 2

    Select Case picDraw.Type
    Case vbPicTypeBitmap
        BitBlt UserControl.hdc, XPos, YPos, XPos + picDraw.Width, YPos + picDraw.Height, dcTemp, 0, 0, vbSrcCopy

    Case vbPicTypeIcon
        ' Create a bitmap and select it into an DC
        ' Draw Icon onto DC
        DrawIconEx UserControl.hdc, XPos, YPos, picDraw.Handle, picDraw.Width, picDraw.Height, 0&, 0&, DI_NORMAL
    End Select

    'Clean up
    GoSub CleanUp
    '------------------------------------------------------------------

DrawText:
    'Since TextOut won't align, and DrawText doesn't work :P,
    'combine both to make something that does! :)
    'Use DrawText to return the text's height, and textout accordingly!
    '------------------------------------------------------------------

    If bUseBevels And bBevel Then
        If bInnerBevel Then
            FormInnerBevel
        Else
            FormOuterBevel
        End If
    End If

    'Set transparent text rendering
    oldTextDrawMode = SetBkMode(UserControl.hdc, TRANSPARENT)

    'Find out the bounds of the usercontrol's rectangle
    result = GetWindowRect(UserControl.hwnd, rt)

    'Asks DrawText to calculate the height of the text, stick it in result
    result = DrawText(UserControl.hdc, sCaption, Len(sCaption), rt, DT_CALCRECT)

    Select Case lhTextPosition
    Case ciLeft
        XPos = 1
        oldTextAlign = SetTextAlign(UserControl.hdc, TA_LEFT)

    Case ciCenter
        XPos = UserControl.ScaleWidth / 2
        oldTextAlign = SetTextAlign(UserControl.hdc, TA_CENTER)

    Case ciRight
        XPos = UserControl.ScaleWidth - 1
        oldTextAlign = SetTextAlign(UserControl.hdc, TA_RIGHT)

    End Select

    Select Case lvTextPosition
    Case cTop
        YPos = 1

    Case cBottom
        YPos = UserControl.ScaleHeight - result - 1

    Case cMiddle
        YPos = UserControl.ScaleHeight / 2 - result / 2

    Case c3Quarters
        YPos = UserControl.ScaleHeight * (3 / 4) - result / 2 - 1

    End Select

    result = TextOut(UserControl.hdc, XPos, YPos, sCaption, Len(sCaption))

    'Put back the old text alignment style
    SetTextAlign UserControl.hdc, oldTextAlign

    'Put back the old text drawing mode
    SetBkMode UserControl.hdc, oldTextDrawMode

    'Ask the control to repaint itself, since i've changed it's looks.
    UserControl.Refresh
    Exit Sub

    'Error handling
    'If we hit an error 91, which will usually mean that picview didn't
    'point to anything, skip blitting image, render text.
    '------------------------------------------------------------------
ErrorHandler:
    If Err.Number = 91 Then GoTo DrawText: GoSub CleanUp: Exit Sub
    MsgBox "Error in Coolbutton UserControl, DrawBevel sub!" & vbCrLf & CStr(Err.Number) & vbCrLf & Err.Description, vbCritical, "Error!"
    GoSub CleanUp
    Exit Sub
    Resume Next


    'Frees objects and memory
    '------------------------------------------------------------------
CleanUp:
    SelectObject dcTemp, bmpOld
    SelectPalette dcTemp, palOld, True
    RealizePalette dcTemp
    DeleteDC dcTemp
    ReleaseDC ByVal 0&, dcDesktop
    DeleteObject palHalfTone
    Return
End Sub

Public Sub ForceRedraw()
    DrawBevel iBevelType
End Sub

Private Sub FormBevelLines(ByVal side As Integer, ByVal wid As Integer, ByVal Color As Long)
    'This is from www.planet-source-code.com's extensive vb code library.

    'Unfortunately, the code would never cut and paste right for me,
    'so i've forgotten the author's name.
    '(obviously someone that likes maths though, not many people use dx.. :) )

    Dim X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer
    Dim rightX As Integer, bottomY As Integer
    Dim dx1 As Integer, dx2 As Integer, dy1 As Integer, dy2 As Integer
    Dim i As Integer

    rightX = UserControl.ScaleWidth - 1
    bottomY = UserControl.ScaleHeight - 1

    Select Case side
    Case 0
        'Left side
        X1 = 0
        dx1 = 1
        X2 = 0
        dx2 = 1
        Y1 = 0
        dy1 = 1
        Y2 = bottomY + 1
        dy2 = -1

    Case 1
        'Right side
        X1 = rightX
        dx1 = -1
        X2 = X1
        dx2 = dx1
        Y1 = 0
        dy1 = 1
        Y2 = bottomY + 1
        dy2 = -1

    Case 2
        'Top side
        X1 = 0
        dx1 = 1
        X2 = rightX
        dx2 = -1
        Y1 = 0
        dy1 = 1
        Y2 = 0
        dy2 = 1

    Case 3
        'Bottom side
        X1 = 1
        dx1 = 1
        X2 = rightX + 1
        dx2 = -1
        Y1 = bottomY
        dy1 = -1
        Y2 = Y1
        dy2 = dy1
    End Select


    For i = 1 To wid

        UserControl.Line (X1, Y1)-(X2, Y2), Color
        X1 = X1 + dx1
        X2 = X2 + dx2
        Y1 = Y1 + dy1
        Y2 = Y2 + dy2

    Next i

End Sub

Private Sub FormOuterBevel()
    UserControl.ScaleMode = vbPixels

    FormBevelLines 0, ciBevelWidth, colour_Highlight
    FormBevelLines 1, ciBevelWidth, colour_LowLight
    FormBevelLines 2, ciBevelWidth, colour_Highlight
    FormBevelLines 3, ciBevelWidth, colour_LowLight
End Sub

Private Sub FormInnerBevel()
    UserControl.ScaleMode = vbPixels

    FormBevelLines 0, ciBevelWidth, colour_LowLight
    FormBevelLines 1, ciBevelWidth, colour_Highlight
    FormBevelLines 2, ciBevelWidth, colour_LowLight
    FormBevelLines 3, ciBevelWidth, colour_Highlight
End Sub

Private Sub FreeDimmedBitmaps()
    If hMouseOverBitmap Then DeleteObject hMouseOverBitmap: hMouseOverBitmap = 0
    If hMouseDownBitmap Then DeleteObject hMouseDownBitmap: hMouseDownBitmap = 0
End Sub

Private Sub GenerateDimmedPictures()
    If picNormal Is Nothing Then Exit Sub

    'Wrote this function half asleep,
    'i hope there's no bugs.

    'This is a really screwy function... :))

    Screen.MousePointer = vbHourglass
    DoEvents

    'Declare variables
    Dim Quads() As RGBQUAD, LongColours() As Long
    Dim result As Long, bmp As BITMAP
    Dim lSize As Long
    Dim i As Long
    Dim hTempDC As Long
    Dim oldBitmap As Long
    Dim bmpinfo As BITMAPINFO
    Dim ti As Integer
    Dim tCol As Long
    Dim srcPtr As Long, dstPtr As Long
    Dim colIgnore As Long

    'VB stores colours in the opposite order of what windows does.
    'which is hell annoying. Alignment and order is different, so
    'i have to rearrange to get it right.

    colIgnore = CLng(colour_Ignore)

    Dim bArray1(3) As Byte
    Dim bArray2(3) As Byte

    srcPtr = VarPtr(colIgnore)
    dstPtr = VarPtr(bArray1(0))

    CopyMemory ByVal dstPtr, ByVal srcPtr, Len(colIgnore)

    bArray2(0) = bArray1(2)
    bArray2(1) = bArray1(1)
    bArray2(2) = bArray1(0)
    bArray2(3) = 0

    srcPtr = VarPtr(bArray2(0))
    dstPtr = VarPtr(colIgnore)

    CopyMemory ByVal dstPtr, ByVal srcPtr, Len(colIgnore)

    'ColIgnore has the colour to ignore in API nice terms.

    'Get the bitmap's dimensions
    result = GetObject(picNormal.Handle, Len(bmp), bmp)

    'Find out the size of the array i need
    lSize = bmp.bmWidth * bmp.bmHeight

    'Make a DC so i can use GetDIBits, SetDIBits
    hTempDC = CreateCompatibleDC(ByVal 0&)

    'Select the bitmap to the DC
    oldBitmap = SelectObject(hTempDC, picNormal.Handle)

    'Alloc mem
    ReDim Quads(lSize)
    ReDim LongColours(lSize)

    'Create info struct, to read raw data in RGB format

    'Asking for the data in RLE format might be a lot faster to
    'process, there's an idea for a speedup.

    With bmpinfo.bmiHeader
        .biSize = Len(bmpinfo.bmiHeader)
        .biWidth = bmp.bmWidth
        .biHeight = bmp.bmHeight
        .biPlanes = bmp.bmPlanes
        .biBitCount = 32
        .biCompression = BI_RGB
    End With

    'Get the data, in Quad and Long form.
    result = GetDIBits(hTempDC, picNormal.Handle, 0&, bmp.bmHeight, Quads(0), bmpinfo, DIB_RGB_COLORS)
    result = GetDIBits(hTempDC, picNormal.Handle, 0&, bmp.bmHeight, LongColours(0), bmpinfo, DIB_RGB_COLORS)

    'Decrease brightness of the bitmap
    For i = LBound(Quads, 1) To UBound(Quads, 1)

        If Not LongColours(i) = colIgnore Then
            With Quads(i)
                .rgbBlue = .rgbBlue * csDimPercent
                .rgbGreen = .rgbGreen * csDimPercent
                .rgbRed = .rgbRed * csDimPercent
            End With
        End If
    Next i

    'Delete any bitmap if already created
    If hMouseDownBitmap Then DeleteObject hMouseDownBitmap

    'Create a bitmap
    hMouseDownBitmap = CreateCompatibleBitmap(UserControl.hdc, bmp.bmWidth, bmp.bmHeight)

    'Select new bitmap
    result = SelectObject(hTempDC, hMouseDownBitmap)

    'Write bits to it
    result = SetDIBits(hTempDC, hMouseDownBitmap, 0, bmp.bmHeight, Quads(0), bmpinfo, DIB_RGB_COLORS)

    'Part 1 done.
    '------------------------------------------------------------------

    'Select original image
    SelectObject hTempDC, picNormal.Handle

    'Get original data again
    result = GetDIBits(hTempDC, picNormal.Handle, 0, bmp.bmHeight, Quads(0), bmpinfo, DIB_RGB_COLORS)

    'Brighten, watching for overflows
    For i = LBound(Quads, 1) To UBound(Quads, 1)

        If Not LongColours(i) = colIgnore Then

            With Quads(i)
                ti = .rgbBlue * csBriPercent
                If ti < cbMaxValue Then
                    .rgbBlue = ti
                Else
                    .rgbBlue = cbMaxValue
                End If

                ti = .rgbGreen * csBriPercent
                If ti < cbMaxValue Then
                    .rgbGreen = ti
                Else
                    .rgbGreen = cbMaxValue
                End If

                ti = .rgbRed * csBriPercent
                If ti < cbMaxValue Then
                    .rgbRed = ti
                Else
                    .rgbRed = cbMaxValue
                End If
            End With
        End If

    Next i

    'Delete old bitmap if present
    If hMouseOverBitmap Then DeleteObject hMouseOverBitmap

    'Create new bitmap
    hMouseOverBitmap = CreateCompatibleBitmap(UserControl.hdc, bmp.bmWidth, bmp.bmHeight)

    'Select bitmap to DC
    SelectObject hTempDC, hMouseOverBitmap

    'Copy data over
    result = SetDIBits(hTempDC, hMouseOverBitmap, 0, bmp.bmHeight, Quads(0), bmpinfo, DIB_RGB_COLORS)

    'Part 2 done
    '------------------------------------------------------------------

    DoEvents

    'Clean up
    '------------------------------------------------------------------

    'Dealloc memory
    Erase Quads()
    Erase LongColours

    'Select back old bitmap
    SelectObject hTempDC, oldBitmap

    'Delete the DC
    result = DeleteDC(hTempDC)

    Screen.MousePointer = vbNormal
End Sub

Private Function HasBackColourProperty(ByVal ctrl As Object) As Boolean
    On Error GoTo ErrorHandler

    Dim colourTemp As OLE_COLOR

    colourTemp = ctrl.BackColor
    HasBackColourProperty = True
    Exit Function

ErrorHandler:
    Exit Function
End Function

Public Property Get Picture() As StdPicture
    Set Picture = picNormal
End Property
Public Property Set Picture(ByVal pNewValue As StdPicture)
    Set picNormal = pNewValue
    PropertyChanged "Picture"

    If bAutoSize Then AutoSizeControl
    If bAutoDim Then GenerateDimmedPictures
    DrawBevel iBevelType
End Property

Public Property Get PictureOver() As StdPicture
    Set PictureOver = picMouseOver
End Property
Public Property Set PictureOver(ByVal pNewValue As StdPicture)
    Set picMouseOver = pNewValue
    PropertyChanged "PictureOver"
End Property

Public Property Get PictureDown() As StdPicture
    Set PictureDown = picMouseDown
End Property
Public Property Set PictureDown(ByVal pNewValue As StdPicture)
    Set picMouseDown = pNewValue
    PropertyChanged "PictureDown"
End Property

Public Property Get colourHighlight() As OLE_COLOR
    colourHighlight = colour_Highlight
End Property
Public Property Let colourHighlight(ByVal cNewValue As OLE_COLOR)
    colour_Highlight = cNewValue
    PropertyChanged "colourHighlight"
End Property

Public Property Get colourLowLight() As OLE_COLOR
    colourLowLight = colour_LowLight
End Property
Public Property Let colourLowLight(ByVal cNewValue As OLE_COLOR)
    colour_LowLight = cNewValue
    PropertyChanged "colourLowLight"
End Property

Public Property Get colourBackColour() As OLE_COLOR
    colourBackColour = colour_BackColour
End Property
Public Property Let colourBackColour(ByVal cNewValue As OLE_COLOR)
    colour_BackColour = cNewValue
    PropertyChanged "colourBackColour"
    UserControl.BackColor = cNewValue
    DrawBevel iBevelType
End Property

Public Property Get colourTextStdColour() As OLE_COLOR
    colourTextStdColour = colour_TextStdColour
End Property
Public Property Let colourTextStdColour(ByVal cNewValue As OLE_COLOR)
    colour_TextStdColour = cNewValue
    PropertyChanged "colourTextStdColour"
    DrawBevel iBevelType
End Property

Public Property Get colourTextOverColour() As OLE_COLOR
    colourTextOverColour = colour_TextOverColour
End Property
Public Property Let colourTextOverColour(ByVal cNewValue As OLE_COLOR)
    colour_TextOverColour = cNewValue
    PropertyChanged "colourTextOverColour"
End Property

Public Property Get Caption() As String
    Caption = sCaption
End Property
Public Property Let Caption(ByVal sNewValue As String)
    sCaption = sNewValue
    PropertyChanged "Caption"

    Dim i As Integer, ts As String

    i = InStr(1, sNewValue, "&", vbBinaryCompare)
    If i <> 0 And i <> Len(sNewValue) Then
        ts = Mid$(sNewValue, i + 1, 1)
        UserControl.AccessKeys = ts
    End If

    DrawBevel iBevelType
End Property

Public Property Get UsePictures() As Boolean
    UsePictures = bUsePictures
End Property
Public Property Let UsePictures(ByVal bNewValue As Boolean)
    bUsePictures = bNewValue
    PropertyChanged "UsePictures"
    DrawBevel iBevelType
End Property

Public Property Get UseBevels() As Boolean
    UseBevels = bUseBevels
End Property
Public Property Let UseBevels(ByVal bNewValue As Boolean)
    bUseBevels = bNewValue
    PropertyChanged "UseBevels"
    DrawBevel iBevelType
End Property

Public Property Get UseDippedControls() As Boolean
    UseDippedControls = bDipControls
End Property
Public Property Let UseDippedControls(ByVal bNewValue As Boolean)
    bDipControls = bNewValue
    PropertyChanged "UseDippedControls"
End Property

Public Property Get AutoSize() As Boolean
    AutoSize = bAutoSize
End Property
Public Property Let AutoSize(ByVal bNewValue As Boolean)
    bAutoSize = bNewValue
    PropertyChanged "AutoSize"
    If bAutoSize Then AutoSizeControl
End Property

Public Property Get UseUnderlineOnFocus() As Boolean
    UseUnderlineOnFocus = bUnderlineFocus
End Property
Public Property Let UseUnderlineOnFocus(ByVal bNewValue As Boolean)
    bUnderlineFocus = bNewValue
    PropertyChanged "UseUnderlineOnFocus"
End Property

Public Property Get CaptionFont() As Font
    Set CaptionFont = UserControl.Font
End Property
Public Property Set CaptionFont(ByVal fNewValue As Font)
    Set UserControl.Font = fNewValue
    PropertyChanged "CaptionFont"
    DrawBevel iBevelType
End Property

Public Property Get Enabled() As Boolean
    Enabled = bEnabled
End Property
Public Property Let Enabled(ByVal bNewValue As Boolean)
    bEnabled = bNewValue
    PropertyChanged "Enabled"
    DrawBevel iBevelType
End Property

Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Let hwnd(ByVal lnewValue As Long)
    'Do nothing - This will be visible in the property box
End Property

Public Property Get AlwaysDrawBevel() As Boolean
    AlwaysDrawBevel = bButtonsAlwaysUp
End Property
Public Property Let AlwaysDrawBevel(ByVal bNewValue As Boolean)
    bButtonsAlwaysUp = bNewValue
    PropertyChanged "AlwaysDrawBevel"
    ForceRedraw
End Property

Public Property Get AutoDim() As Boolean
    AutoDim = bAutoDim
End Property
Public Property Let AutoDim(ByVal bNewValue As Boolean)
    bAutoDim = bNewValue
    PropertyChanged "AutoDim"
    If bAutoDim Then
        If Ambient.UserMode Then GenerateDimmedPictures
    Else
        FreeDimmedBitmaps
    End If
End Property

Public Property Get TextPositionV() As eVTextPosition
    TextPositionV = lvTextPosition
End Property
Public Property Let TextPositionV(ByVal iNewValue As eVTextPosition)
    lvTextPosition = iNewValue
    PropertyChanged "TextPositionV"
    DrawBevel iBevelType
End Property

Public Property Get TextPositionH() As eHTextPosition
    TextPositionH = lhTextPosition
End Property
Public Property Let TextPositionH(ByVal iNewValue As eHTextPosition)
    lhTextPosition = iNewValue
    PropertyChanged "TextPositionH"
    DrawBevel iBevelType
End Property

Public Property Get colourIgnore() As OLE_COLOR
    colourIgnore = colour_Ignore
End Property
Public Property Let colourIgnore(ByVal cNewValue As OLE_COLOR)
    colour_Ignore = cNewValue
    PropertyChanged "colourIgnore"
End Property

Public Property Get AutoColour() As Boolean
    AutoColour = bAutoColour
End Property
Public Property Let AutoColour(ByVal bNewValue As Boolean)

    Static bUsingOldColour As Boolean
    Static colourOld As OLE_COLOR

    If HasBackColourProperty(UserControl.Extender.Container) Then
        If bNewValue Then
            colourOld = colourBackColour
            colourBackColour = UserControl.Extender.Container.BackColor
            bUsingOldColour = True
        Else
            If bUsingOldColour Then colourBackColour = colourOld
        End If

        bAutoColour = bNewValue
        PropertyChanged "AutoColour"
    Else
        bNewValue = False
        bAutoColour = False
'        VBA.MsgBox "Sorry, AutoColour can't be changed, because the container doesn't support a BackColor property!", vbExclamation
    End If
End Property
