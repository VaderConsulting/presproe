VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTooltips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const WIN32_IE = &H400

Private hWndTT As Long
Private m_Enabled As Boolean
Private m_Icon As InfoTitleConstants
Private m_Title As String

Public Enum DelayTimeConstants
    dtAutoPop = &H2
    dtInitial = &H3
    dtReshow = &H1
End Enum

Public Enum SetDelayTimeConstants
    sdtAutoPop = &H2
    sdtInitial = &H3
    sdtReshow = &H1
    sdtAutomatic = &H0
End Enum

Public Enum InfoTitleConstants
    itNoIcon = 0
    itInfoIcon = 1
    itWarningIcon = 2
    itErrorIcon = 3
End Enum

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TOOLINFO
    cbSize As Long
    uFlags As Long
    hwnd As Long
    uId As Long
    WinRect As RECT
    hInst As Long
    lpszText As String
    #If WIN32_IE >= &H300 Then
        lParam As Long
    #End If
End Type

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Const WS_EX_TOPMOST = &H8&
Private Const TOOLTIPS_CLASS = "tooltips_class32"
Private Const WS_POPUP = &H80000000
Private Const CW_USEDEFAULT = &H80000000

Private Declare Function InitCommonControlsEx Lib "comctl32" (lpInitCtrls As INITCOMMONCONTROLEXSTRUCT) As Long
Private Const ICC_WIN95_CLASSES = &HFF
Private Const ICC_BAR_CLASSES = &H4

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
#If Not UNICODE Then
    Private Const TTM_ADDTOOL = WM_USER + 4
    Private Const TTM_DELTOOL = WM_USER + 5
    Private Const TTM_ENUMTOOLS = WM_USER + 14
    Private Const TTM_HITTEST = WM_USER + 10
    Private Const TTM_NEWTOOLRECT = WM_USER + 6
    Private Const TTM_SETTITLE = WM_USER + 32
    Private Const TTM_UPDATETIPTEXT = WM_USER + 12
#Else
    Private Const TTM_ADDTOOL = WM_USER + 50
    Private Const TTM_DELTOOL = WM_USER + 51
    Private Const TTM_ENUMTOOLS = WM_USER + 58
    Private Const TTM_HITTEST = WM_USER + 55
    Private Const TTM_NEWTOOLRECT = WM_USER + 52
    Private Const TTM_SETTITLE = WM_USER + 33
    Private Const TTM_UPDATETIPTEXT = WM_USER + 57
#End If
Private Const TTM_ACTIVATE = WM_USER + 1
Private Const TTM_ADJUSTRECT = WM_USER + 31
Private Const TTM_GETTOOLCOUNT = WM_USER + 13
Private Const TTM_GETBUBBLESIZE = WM_USER + 30
Private Const TTM_GETDELAYTIME = WM_USER + 21
Private Const TTM_GETMARGIN = WM_USER + 27
Private Const TTM_GETMAXTIPWIDTH = WM_USER + 25
Private Const TTM_GETTIPBKCOLOR = WM_USER + 22
Private Const TTM_GETTIPTEXTCOLOR = WM_USER + 23
Private Const TTM_POP = WM_USER + 28
Private Const TTM_SETDELAYTIME = WM_USER + 3
Private Const TTM_SETMARGIN = WM_USER + 26
Private Const TTM_SETMAXTIPWIDTH = WM_USER + 24
Private Const TTM_SETTIPBKCOLOR = WM_USER + 19
Private Const TTM_SETTIPTEXTCOLOR = WM_USER + 20
Private Const TTM_TRACKACTIVATE = WM_USER + 17
Private Const TTM_TRACKPOSITION = WM_USER + 18

Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Const TTF_IDISHWND = &H1
Private Const TTF_SUBCLASS = &H10

Private Type Size
    cx As Long
    cy As Long
End Type
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hDC As Long, ByVal lpString As String, ByVal cbString As Long, lpSize As Size) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Enum TooltipFlagConstants
    ttfAlwaysTip = &H1
    ttfNoPrefix = &H2
    ttfNoAnimate = &H10
    ttfNoFade = &H20
    ttfBalloon = &H40
End Enum

Public Enum ToolFlagConstants
    tfCenterTip = &H2
    tfRtlReading = &H4
    tfTrack = &H20
    tfAbsolute = &H80
    tfTransparent = &H100
End Enum

Private Type INITCOMMONCONTROLEXSTRUCT
    dwSize As Long
    dwICC As Long
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type TTHITTESTINFO
    hwnd As Long
    pt As POINTAPI
    ti As TOOLINFO
End Type

Private Type OLECOLOR
    RedOrSys As Byte
    Green As Byte
    Blue As Byte
    Type As Byte
End Type
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Sub AddTool(Ctl As Control, Flags As ToolFlagConstants, Optional Text As String)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info.cbSize = Len(Info)
    Info.uFlags = Flags
    If Not (Flags And tfTrack) Then
        Info.uFlags = Info.uFlags Or TTF_SUBCLASS
    End If
    Info.uFlags = Info.uFlags Or TTF_IDISHWND
    Info.hwnd = Ctl.Container.hwnd
    Info.hInst = App.hInstance
    Info.uId = Ctl.hwnd
    If Len(Text) > 0 Then
        Info.lpszText = Text
    End If
    
    SendMessage hWndTT, TTM_ADDTOOL, 0, Info
End Sub
Public Sub Create(Frm As Form, Flags As TooltipFlagConstants)
    Dim InitCtrls As INITCOMMONCONTROLEXSTRUCT
    
    InitCtrls.dwSize = Len(InitCtrls)
    InitCtrls.dwICC = ICC_WIN95_CLASSES Or ICC_BAR_CLASSES
    
    Class_Terminate
    m_Enabled = True
    m_Icon = itNoIcon
    m_Title = ""
    
    InitCommonControlsEx InitCtrls
    
    hWndTT = CreateWindowEx(WS_EX_TOPMOST, TOOLTIPS_CLASS, "", _
        WS_POPUP Or Flags, CW_USEDEFAULT, CW_USEDEFAULT, _
        CW_USEDEFAULT, CW_USEDEFAULT, Frm.hwnd, 0, App.hInstance, _
        ByVal 0&)
    
    SetWindowPos hWndTT, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or _
        SWP_NOSIZE Or SWP_NOACTIVATE
End Sub


Public Sub DeleteTool(Ctl As Control)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info = GetToolInfo(Ctl.hwnd)
    
    If Info.cbSize <> 0 Then
        SendMessage hWndTT, TTM_DELTOOL, 0, Info
    End If
End Sub

Public Property Get Enabled() As Boolean
    If hWndTT = 0 Then Exit Property
    
    Enabled = m_Enabled
End Property

Public Property Let Enabled(NewState As Boolean)
    m_Enabled = NewState
    
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_ACTIVATE, m_Enabled, ByVal 0&
End Property

Public Function GetDelayTime(TimeMode As DelayTimeConstants) As Long
    If hWndTT = 0 Then Exit Function
    
    GetDelayTime = SendMessage(hWndTT, TTM_GETDELAYTIME, TimeMode, _
        0&)
End Function

Public Property Get Icon() As InfoTitleConstants
    If hWndTT = 0 Then Exit Property
    
    Icon = m_Icon
End Property

Public Function MakeLong(Int1 As Integer, Int2 As Integer) As Long
    Dim Ints(1 To 2) As Integer
    
    CopyMemory Ints(1), Int1, Len(Int1)
    CopyMemory Ints(2), Int2, Len(Int2)
    
    CopyMemory MakeLong, Ints(1), Len(MakeLong)
End Function

Public Sub SetTracking(Ctl As Control, Tracking As Boolean)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info = GetToolInfo(Ctl.hwnd)
    If Info.cbSize <> 0 Then
        SendMessage hWndTT, TTM_TRACKACTIVATE, Tracking, Info
    End If
End Sub

Public Sub SetTrackPosition(Ctl As Control, X As Integer, Y As Integer)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info = GetToolInfo(Ctl.hwnd)
    If Info.cbSize <> 0 Then
        SendMessage hWndTT, TTM_TRACKPOSITION, 0, _
            ByVal MakeLong(X, Y)
    End If
End Sub

Public Property Get Title() As String
    If hWndTT = 0 Then Exit Property
    
    Title = m_Title
End Property

Public Property Let Icon(NewIcon As InfoTitleConstants)
    If hWndTT = 0 Then Exit Property
    
    m_Icon = NewIcon
    
    SendMessage hWndTT, TTM_SETTITLE, m_Icon, ByVal Title
End Property
Public Property Let Title(NewTitle As String)
    If hWndTT = 0 Then Exit Property
    
    m_Title = NewTitle
    
    SendMessage hWndTT, TTM_SETTITLE, Icon, ByVal m_Title
End Property

Public Sub SetDelayTime(TimeMode As SetDelayTimeConstants, Time As Long)
    If hWndTT = 0 Then Exit Sub
    
    SendMessage hWndTT, TTM_SETDELAYTIME, TimeMode, ByVal Time
End Sub

Public Sub HideTips()
    If hWndTT = 0 Then Exit Sub
    
    SendMessage hWndTT, TTM_POP, 0, ByVal 0&
End Sub

Public Function HitTest(Ctl As Control, X As Long, Y As Long) As Boolean
    Dim Info As TTHITTESTINFO
    
    If hWndTT = 0 Then Exit Function
    
    With Info
        .hwnd = Ctl.hwnd
        .pt.X = X
        .pt.Y = Y
        .ti.cbSize = Len(.ti)
    End With
    
    HitTest = (SendMessage(Ctl.hwnd, TTM_HITTEST, 0, Info) <> 0)
End Function

Public Property Get MarginLeft() As Long
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_GETMARGIN, 0, Rct
    
    MarginLeft = Rct.Left
End Property

Public Property Let MarginLeft(NewMargin As Long)
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    With Rct
        .Left = NewMargin
        .Top = MarginTop
        .Right = MarginRight
        .Bottom = MarginBottom
    End With
    
    SendMessage hWndTT, TTM_SETMARGIN, 0, Rct
End Property
Public Property Let MarginTop(NewMargin As Long)
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    With Rct
        .Left = MarginLeft
        .Top = NewMargin
        .Right = MarginRight
        .Bottom = MarginBottom
    End With
    
    SendMessage hWndTT, TTM_SETMARGIN, 0, Rct
End Property

Public Property Let MarginRight(NewMargin As Long)
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    With Rct
        .Left = MarginLeft
        .Top = MarginTop
        .Right = NewMargin
        .Bottom = MarginBottom
    End With
    
    SendMessage hWndTT, TTM_SETMARGIN, 0, Rct
End Property

Public Property Let MarginBottom(NewMargin As Long)
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    With Rct
        .Left = MarginLeft
        .Top = MarginTop
        .Right = MarginRight
        .Bottom = NewMargin
    End With
    
    SendMessage hWndTT, TTM_SETMARGIN, 0, Rct
End Property

Public Property Get MarginTop() As Long
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_GETMARGIN, 0, Rct
    
    MarginTop = Rct.Top
End Property

Public Property Get MarginRight() As Long
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_GETMARGIN, 0, Rct
    
    MarginRight = Rct.Right
End Property

Public Property Get MarginBottom() As Long
    Dim Rct As RECT
    
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_GETMARGIN, 0, Rct
    
    MarginBottom = Rct.Bottom
End Property

Public Function GetShellVersion() As Long

End Function

Private Function GetToolInfo(ToolWnd As Long) As TOOLINFO
    Dim n As Long, Info As TOOLINFO, Found As Boolean
    
    If hWndTT = 0 Then Exit Function
    
    Info.cbSize = Len(Info)
    
    For n = 0 To ToolCount - 1
        If SendMessage(hWndTT, TTM_ENUMTOOLS, n, Info) <> 0 Then
            If Info.uId = ToolWnd Then
                Found = True
                Exit For
            End If
        End If
    Next n
    
    If Found Then GetToolInfo = Info
End Function

Public Function GetTooltipWidth(Ctl As Control) As Long
    Dim TipSize As Long, Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Function
    
    Info = GetToolInfo(Ctl.hwnd)
    
    If Info.cbSize <> 0 Then
        TipSize = SendMessage(hWndTT, TTM_GETBUBBLESIZE, 0, _
            Info)
    End If
    
    GetTooltipWidth = LoWord(TipSize)
End Function

Public Function GetTooltipHeight(Ctl As Control) As Long
    Dim TipSize As Long, Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Function
    
    Info = GetToolInfo(Ctl.hwnd)
    
    If Info.cbSize <> 0 Then
        TipSize = SendMessage(hWndTT, TTM_GETBUBBLESIZE, 0, _
            Info)
    End If
    
    GetTooltipHeight = HiWord(TipSize)
End Function

Private Function LoWord(DWord As Long) As Integer
    Dim Ints(1 To 2) As Integer
    
    CopyMemory Ints(1), DWord, Len(DWord)
    
    LoWord = Ints(1)
End Function

Private Function HiWord(DWord As Long) As Integer
    Dim Ints(1 To 2) As Integer
    
    CopyMemory Ints(1), DWord, Len(DWord)
    
    HiWord = Ints(2)
End Function

Public Property Get MaxTipWidth() As Long
    If hWndTT = 0 Then Exit Property
    
    MaxTipWidth = SendMessage(hWndTT, TTM_GETMAXTIPWIDTH, 0, ByVal 0&)
End Property

Public Property Let MaxTipWidth(NewWidth As Long)
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_SETMAXTIPWIDTH, 0, ByVal NewWidth
End Property

Public Sub SetToolRect(Ctl As Control, Left As Long, Top As Long, Right As Long, Bottom As Long)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info = GetToolInfo(Ctl.hwnd)
    With Info.WinRect
        .Left = Left
        .Top = Top
        .Right = Right
        .Bottom = Bottom
    End With
    
    If Info.cbSize <> 0 Then
        SendMessage hWndTT, TTM_NEWTOOLRECT, 0, Info
    End If
End Sub

Public Property Get ToolCount() As Long
    If hWndTT = 0 Then Exit Sub
    
    ToolCount = SendMessage(hWndTT, TTM_GETTOOLCOUNT, 0, ByVal 0&)
End Property

Public Property Get BackColor() As OLE_COLOR
    If hWndTT = 0 Then Exit Property
    
    BackColor = SendMessage(hWndTT, TTM_GETTIPBKCOLOR, 0, ByVal 0&)
End Property
Public Property Let BackColor(NewColor As OLE_COLOR)
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_SETTIPBKCOLOR, WinColor(NewColor), ByVal 0&
End Property

Public Sub SetTipText(Ctl As Control, NewText As String)
    Dim Info As TOOLINFO
    
    If hWndTT = 0 Then Exit Sub
    
    Info = GetToolInfo(Ctl.hwnd)
    If Info.cbSize <> 0 Then
        Info.lpszText = NewText
        Info.cbSize = Len(Info)
        SendMessage hWndTT, TTM_UPDATETIPTEXT, 0, Info
    End If
End Sub

Public Property Let ForeColor(NewColor As OLE_COLOR)
    If hWndTT = 0 Then Exit Property
    
    SendMessage hWndTT, TTM_SETTIPTEXTCOLOR, WinColor(NewColor), ByVal 0&
End Property

Private Function WinColor(VBColor As Long) As Long
    Dim SysClr As OLECOLOR
    
    CopyMemory SysClr, VBColor, Len(SysClr)
    
    If SysClr.Type = &H80 Then 'It is a system color
        'SysClr.RedOrSys is the index of the system color
        WinColor = GetSysColor(SysClr.RedOrSys)
    Else
        WinColor = VBColor
    End If
End Function
Public Property Get ForeColor() As OLE_COLOR
    If hWndTT = 0 Then Exit Property
    
    ForeColor = SendMessage(hWndTT, TTM_GETTIPTEXTCOLOR, 0, ByVal 0&)
End Property

Public Function ToolTipWidthFromText(hDC As Long, Text As String) As Long
    Dim Rct As RECT, Siz As Size
    
    If hWndTT = 0 Then Exit Function
    
    GetTextExtentPoint32 hDC, Text, LenB(Text), Siz
    
    Rct.Right = Siz.cx
    Rct.Bottom = Siz.cy
    SendMessage hWndTT, TTM_ADJUSTRECT, True, Rct
    
    ToolTipWidthFromText = (Rct.Right - Rct.Left)
End Function


Public Function ToolTipHeightFromText(hDC As Long, Text As String) As Long
    Dim Rct As RECT, Siz As Size
    
    If hWndTT = 0 Then Exit Function
    
    GetTextExtentPoint32 hDC, Text, LenB(Text), Siz
    
    Rct.Right = Siz.cx
    Rct.Bottom = Siz.cy
    SendMessage hWndTT, TTM_ADJUSTRECT, True, Rct
    
    ToolTipHeightFromText = (Rct.Bottom - Rct.Top)
End Function
Public Function TextHeightFromToolTipHeight(hDC As Long, TextHeight As Long) As Long
    Dim Rct As RECT, Siz As Size
    
    If hWndTT = 0 Then Exit Function
    
    Rct.Bottom = TextHeight
    SendMessage hWndTT, TTM_ADJUSTRECT, False, Rct
    
    TextHeightFromToolTipHeight = (Rct.Bottom - Rct.Top)
End Function

Public Function TextWidthFromToolTipWidth(hDC As Long, TextWidth As Long) As Long
    Dim Rct As RECT, Siz As Size
    
    If hWndTT = 0 Then Exit Function
    
    Rct.Right = TextWidth
    SendMessage hWndTT, TTM_ADJUSTRECT, False, Rct
    
    TextWidthFromToolTipWidth = (Rct.Right - Rct.Left)
End Function

Private Sub Class_Terminate()
    If hWndTT <> 0 Then DestroyWindow hWndTT
End Sub


