VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1275
   ScaleWidth      =   1500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1740
      Top             =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal HDC As Long, ByVal crColor As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function StretchDIBits& Lib "gdi32" (ByVal HDC&, ByVal x&, ByVal y&, ByVal dX&, ByVal dy&, ByVal SrcX&, ByVal SrcY&, ByVal Srcdx&, ByVal Srcdy&, Bits As Any, BInf As Any, ByVal Usage&, ByVal Rop&)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal HDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Boolean, ByVal fdwUnderline As Boolean, ByVal fdwStrikeOut As Boolean, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal HDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal HDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private SCRw&
Private SCRh&
Private CentreW&
Private CentreH&
Private ClockSize&
Private m_oLine As LineGS
Private mBlank As cDIBSection
Private mBuffer As cDIBSection

Private SSAngle As Single

Const FW_DONTCARE = 0
Const FW_THIN = 100
Const FW_EXTRALIGHT = 200
Const FW_LIGHT = 300
Const FW_NORMAL = 400
Const FW_MEDIUM = 500
Const FW_SEMIBOLD = 600
Const FW_BOLD = 700
Const FW_EXTRABOLD = 800
Const FW_HEAVY = 900
Const FW_BLACK = FW_HEAVY
Const FW_DEMIBOLD = FW_SEMIBOLD
Const FW_REGULAR = FW_NORMAL
Const FW_ULTRABOLD = FW_EXTRABOLD
Const FW_ULTRALIGHT = FW_EXTRALIGHT
'used with fdwCharSet
Const ANSI_CHARSET = 0
Const DEFAULT_CHARSET = 1
Const SYMBOL_CHARSET = 2
Const SHIFTJIS_CHARSET = 128
Const HANGEUL_CHARSET = 129
Const CHINESEBIG5_CHARSET = 136
Const OEM_CHARSET = 255
'used with fdwOutputPrecision
Const OUT_CHARACTER_PRECIS = 2
Const OUT_DEFAULT_PRECIS = 0
Const OUT_DEVICE_PRECIS = 5
'used with fdwClipPrecision
Const CLIP_DEFAULT_PRECIS = 0
Const CLIP_CHARACTER_PRECIS = 1
Const CLIP_STROKE_PRECIS = 2
'used with fdwQuality
Const DEFAULT_QUALITY = 0
Const DRAFT_QUALITY = 1
Const PROOF_QUALITY = 2
'used with fdwPitchAndFamily
Const DEFAULT_PITCH = 0
Const FIXED_PITCH = 1
Const VARIABLE_PITCH = 2
Const OPAQUE = 2
Const TRANSPARENT = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40
Const LOGPIXELSY = 90
Const COLOR_WINDOW = 5
Private Function CreateMyFont(nSize&, sFontFace$, bBold As Boolean, bItalic As Boolean) As Long
Static r&, d&

    DeleteDC r: r = GetDC(0)
    d = GetDeviceCaps(r, LOGPIXELSY)
    CreateMyFont = CreateFont(-MulDiv(nSize, d, 72), 0, 0, 0, _
                              IIf(bBold, FW_BOLD, FW_NORMAL), bItalic, False, False, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, _
                              CLIP_DEFAULT_PRECIS, PROOF_QUALITY, DEFAULT_PITCH, sFontFace) 'gdi 2
End Function

Private Sub SetFont(DC&, sFace$, nSize&)
Static c&
    ReleaseDC DC, c: DeleteDC c
    c = CreateMyFont(nSize, sFace, False, False)
    DeleteObject SelectObject(DC, c)
End Sub
Private Sub Form_DblClick()
    Set mBlank = Nothing
    Set mBuffer = Nothing
    Set m_oLine = Nothing
    Unload Me
    End
End Sub
Sub SetAlpha()
    Dim Ret&
    Static AlphaOn As Boolean
    Const LWA_COLORKEY = &H1
    Const LWA_ALPHA = &H2
    Const GWL_EXSTYLE = (-20)
    Const WS_EX_LAYERED = &H80000
    
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    
    SetLayeredWindowAttributes Me.hwnd, 0, IIf(AlphaOn, 255, 96), LWA_ALPHA
    AlphaOn = Not AlphaOn
    SaveSetting "MTCLOCK", "Settings", "opacity", IIf(AlphaOn, "1", "0")
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then CentreClock
If KeyCode = 38 Then ClockSizeChange 1 'up
If KeyCode = 40 Then ClockSizeChange -1 'up
If KeyCode = 84 Then SetWindowTopmost Me.hwnd
If KeyCode = 79 Then SetAlpha
End Sub

Sub ClockSizeChange(nSize%)
    If nSize < 0 And Me.Width <= 450 Then Exit Sub
    If nSize > 0 And Me.Height > (Screen.Height * 0.8) Then Exit Sub
    
    Me.Height = Me.Height + (nSize * 150)
    Me.Width = Me.Height
    SetForm
End Sub
Sub SetForm()
Dim h&
    If Me.Height <> Me.Width Then
        If Me.Height > Me.Width Then
            Me.Height = Me.Width
        Else
            Me.Width = Me.Height
        End If
    End If

    'h = CreateRoundRectRgn(0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY, 21, 21)
    h = CreateEllipticRgn(0, 0, Me.Width \ Screen.TwipsPerPixelX, Me.Height \ Screen.TwipsPerPixelY)
                            
    SetWindowRgn Me.hwnd, h, False
End Sub
Private Sub Form_Load()


    Set m_oLine = New LineGS
    Set mBlank = New cDIBSection
    Set mBuffer = New cDIBSection

    Me.Left = GetSetting("MTCLOCK", "Settings", "left", 0)
    Me.Top = GetSetting("MTCLOCK", "Settings", "top", 0)
    Me.Width = GetSetting("MTCLOCK", "Settings", "width", 1200)
    If GetSetting("MTCLOCK", "Settings", "topmost", "0") = "1" Then SetWindowTopmost Me.hwnd
    If GetSetting("MTCLOCK", "Settings", "opacity", "0") = "1" Then SetAlpha
    
    Me.Height = Me.Width
    
    SetForm


    If Me.Top < 0 Or Me.Top > Screen.Height Or Me.Left < 0 Or Me.Left > Screen.Width Then CentreClock
    
    Timer1.Enabled = True
    
End Sub
Sub CentreClock()
Me.Move Screen.Width \ 2, Screen.Height \ 2, 1200, 1200
SetForm
End Sub
Private Sub SplitRGB(ByVal clr&, r&, g&, b&)
    r = clr And &HFF: g = (clr \ &H100&) And &HFF: b = (clr \ &H10000) And &HFF
End Sub
Private Sub Gradient(DC&, x&, y&, dX&, dy&, ByVal c1&, ByVal c2&, v As Boolean)
Dim r1&, g1&, b1&, r2&, g2&, b2&, b() As Byte
Dim i&, lR!, lG!, lB!, dR!, dG!, dB!, BI&(9), xx&, yy&, dd&, hRPen&
    If dX = 0 Or dy = 0 Then Exit Sub
    If v Then xx = 1: yy = dy: dd = dy Else xx = dX: yy = 1: dd = dX
    SplitRGB c1, r1, g1, b1: SplitRGB c2, r2, g2, b2: ReDim b(dd * 4 - 1)
    dR = (r2 - r1) / (dd - 1): lR = r1: dG = (g2 - g1) / (dd - 1): lG = g1: dB = (b2 - b1) / (dd - 1): lB = b1
    For i = 0 To (dd - 1) * 4 Step 4: b(i + 2) = lR: lR = lR + dR: b(i + 1) = lG: lG = lG + dG: b(i) = lB: lB = lB + dB: Next
    BI(0) = 40: BI(1) = xx: BI(2) = -yy: BI(3) = 2097153: StretchDIBits DC, x, y, dX, dy, 0, 0, xx, yy, b(0), BI(0), 0, vbSrcCopy
End Sub

Function GimmeX(ByVal aIn As Single, lIn As Long) As Long
    GimmeX = Sin(aIn * 0.01745329251994) * lIn
End Function
Function GimmeY(ByVal aIn As Single, lIn As Long) As Long
    GimmeY = Cos(aIn * 0.01745329251994) * lIn
End Function
Function CorrectForAngle(aIn As Single) As Single
CorrectForAngle = 180 - aIn
If CorrectForAngle > 359 Then CorrectForAngle = CorrectForAngle - 360
If CorrectForAngle < 0 Then CorrectForAngle = CorrectForAngle + 360
End Function
Sub DrawHands(the_HDC&)
Dim aMM As Single
Dim aHH As Single
Dim aSS As Single
Dim posX&(1), posY&(1)
Dim hhSize&
Dim mmSize&
Dim ssSize&


    hhSize = ClockSize * 0.5
    mmSize = ClockSize * 0.8
    ssSize = ClockSize

    aHH = CorrectForAngle((Format(Now, "HH") * 30) + (Format(Now, "NN") / 2))
    aMM = CorrectForAngle(Format(Now, "NN") * 6)
    aSS = CorrectForAngle(Format(Now, "ss") * 6)
    
    'hh
    posX(0) = GimmeX(aHH, 5) + CentreW: posY(0) = GimmeY(aHH, 5) + CentreH
    posX(1) = GimmeX(aHH, hhSize) + CentreW: posY(1) = GimmeY(aHH, hhSize) + CentreH
    m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), 0
    
    'mm
    posX(0) = GimmeX(aMM, 5) + CentreW: posY(0) = GimmeY(aMM, 5) + CentreH
    posX(1) = GimmeX(aMM, mmSize) + CentreW: posY(1) = GimmeY(aMM, mmSize) + CentreH
    m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), 0

    'ss
    posX(0) = GimmeX(aSS, 5) + CentreW: posY(0) = GimmeY(aSS, 5) + CentreH
    posX(1) = GimmeX(aSS, ssSize) + CentreW: posY(1) = GimmeY(aSS, ssSize) + CentreH
    m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), RGB(110, 110, 110)
    
    
End Sub
Private Sub SetWindowTopmost(hwnd&)
Static OnTop As Boolean

    If OnTop Then
        SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Else
        SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    End If
    
    OnTop = Not OnTop
    
    SaveSetting "MTCLOCK", "Settings", "topmost", IIf(OnTop, "1", "0")
    
End Sub
Sub DrawClockFace(the_HDC&)
Dim a As Single, b As Single
Dim posX&(1), posY&(1)

    For a = 0 To (ClockSize * 1.5) Step 6
        m_oLine.CircleGP the_HDC, CentreW, CentreH, CLng(a), CLng(a), RGB(100, 150, 190)
    Next
    For a = 0 To 359 Step 6
        posX(0) = GimmeX(a, ClockSize * 0.9) + CentreW
        posY(0) = GimmeY(a, ClockSize * 0.9) + CentreH
        
        posX(1) = GimmeX(a, ClockSize * 0.95) + CentreW
        posY(1) = GimmeY(a, ClockSize * 0.95) + CentreH
        m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), RGB(100, 150, 190)
    Next
    For b = 0 To 5 Step 0.5
        For a = b To 359 Step 30
            posX(0) = GimmeX(a, ClockSize * 0.9) + CentreW
            posY(0) = GimmeY(a, ClockSize * 0.9) + CentreH
            
            posX(1) = GimmeX(a, ClockSize) + CentreW
            posY(1) = GimmeY(a, ClockSize) + CentreH
            m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), 0
        Next
    Next
    For a = 3 To 359 Step 30
        posX(0) = GimmeX(a, ClockSize * 0.9) + CentreW
        posY(0) = GimmeY(a, ClockSize * 0.9) + CentreH
        
        posX(1) = GimmeX(a, ClockSize) + CentreW
        posY(1) = GimmeY(a, ClockSize) + CentreH
        m_oLine.LineGP the_HDC, posX(0), posY(0), posX(1), posY(1), vbWhite
    Next
    
    m_oLine.CircleGP the_HDC, CentreW, CentreH, 4, 4, 0, Thick
    m_oLine.CircleGP the_HDC, CentreW, CentreH, 2, 2, 0
    m_oLine.CircleGP the_HDC, CentreW, CentreH, 3, 3, vbWhite
    
    For b = 0 To 2
        m_oLine.CircleGP the_HDC, CentreW, CentreH, ((Me.Width \ Screen.TwipsPerPixelX) \ 2) - b, ((Me.Height \ Screen.TwipsPerPixelY) \ 2) - b, 0, Thick
    Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbLeftButton Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End If
End Sub

Private Sub Form_Resize()

    SCRw = Me.Width / Screen.TwipsPerPixelX
    SCRh = Me.Height / Screen.TwipsPerPixelY
    CentreW = (SCRw / 2) - 1
    CentreH = (SCRh / 2) - 1
    ClockSize = IIf(CentreW > CentreH, CentreH, CentreW) * 0.9
            
    mBlank.ClearUp
    mBlank.Create SCRw, SCRh
    SetBkMode mBlank.HDC, TRANSPARENT
    Gradient mBlank.HDC, 0, 0, SCRw, SCRh, RGB(100, 150, 190), RGB(200, 210, 240), True
    
    DrawClockFace mBlank.HDC
            
    mBuffer.ClearUp
    mBuffer.Create SCRw, SCRh
    SetBkMode mBuffer.HDC, TRANSPARENT
    
    SetFont mBuffer.HDC, "Small Caps", 7

End Sub
Private Sub BlankToBuffer()
    BitBlt mBuffer.HDC, 0, 0, SCRw, SCRh, mBlank.HDC, 0, 0, vbSrcCopy
End Sub
Private Sub BufferToScreen()
    BitBlt Me.HDC, 0, 0, SCRw, SCRh, mBuffer.HDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "MTCLOCK", "Settings", "left", Me.Left
    SaveSetting "MTCLOCK", "Settings", "top", Me.Top
    SaveSetting "MTCLOCK", "Settings", "width", Me.Width

End Sub

Private Sub Timer1_Timer()
Dim s$
    s = Format(Now, "HH:NN:SS")
    BlankToBuffer
    
    SetTextColor mBuffer.HDC, vbWhite
    TextOut mBuffer.HDC, CentreW - 16, CentreH + 12, s, 8
    SetTextColor mBuffer.HDC, RGB(50, 50, 50)
    TextOut mBuffer.HDC, CentreW - 16, CentreH + 11, s, 8
    
    DrawHands mBuffer.HDC
    BufferToScreen
End Sub


