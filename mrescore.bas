Attribute VB_Name = "mResCore"
Option Explicit

Private Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Any, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As String, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenPtr Lib "kernel32" Alias "lstrlenA" (ByVal lpLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const IdxFile = "__index.txt"

Private ResTypes As RESTYPEOPTIONS
Private TriggerObj As Object
Private DestPath As String
Private CurFile As String
Private CancelResSearch As Long

Private Const RT_BITMAP = 2&
Private Const RT_ICON = 3&
Private Const RT_CURSOR = 1&
Private Const RT_ANICURSOR = 21&
Private Const RT_HTML = 23&
Private Const RT_GROUP_ICON = 14&
Private Const RT_GROUP_CURSOR = 12&
Private Const DIFFERENCE = 11&

Public Type RESTYPEOPTIONS
    Bitmap As Boolean
    Icon As Boolean
    Cursor As Boolean
    AniCursor As Boolean
    Video As Boolean
    HTML As Boolean
End Type

Private Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Const LOAD_LIBRARY_AS_DATAFILE = &H2

Public Enum RESOURCETYPES
    rcrtBitmap
    rcrtIcon
    rcrtCursor
    rcrtAniCursor
    rcrtVideo
    rcrtHTML
End Enum

Private Declare Function FindResource Lib "kernel32" Alias "FindResourceA" (ByVal hInstance As Long, ByVal lpName As String, lpType As Any) As Long
Private Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function LockResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function SizeofResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Private Declare Function PathFindFileName Lib "shlwapi" Alias "PathFindFileNameA" (ByVal pPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal lpszPath As String) As Long
Private Declare Function FreeResource Lib "kernel32" (ByVal hResData As Long) As Long
Private Declare Function PathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecA" (ByVal pszPath As String) As Long
Private Const GCT_LFNCHAR = &H1
Private Declare Function PathGetCharType Lib "shlwapi" Alias "PathGetCharTypeA" (ByVal ch As Byte) As Long
Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
Private Declare Sub PathRemoveExtension Lib "shlwapi" Alias "PathRemoveExtensionA" (ByVal pszPath As String)

Private Type GRPICONDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer
    dwBytesInRes As Long
    nID As Integer
End Type
Private Type ICONDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wPlanes As Integer
    wBitCount As Integer
    dwBytesInRes As Long
    dwImageOffset As Long
End Type
Private Type GRPCURSORDIRENTRY
    wWidth As Integer
    wHeight As Integer
    wPlanes As Integer
    wBitCount As Integer
    lBytesInRes As Long
    wNameOrdinal As Integer
End Type
Private Type CURSORDIRENTRY
    bWidth As Byte
    bHeight As Byte
    bColorCount As Byte
    bReserved As Byte
    wHotspotX As Integer
    wHotspotY As Integer
    dwBytesInRes As Long
    dwImageOffset As Long
End Type
Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type
Private Type BITMAPINFOHEADER
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
Private Type BYTEARRAY
    Data() As Byte
End Type
Private Function AdjustPath(MyPath As String) As String
    If Right(MyPath, 1) <> "\" Then
        AdjustPath = MyPath & "\"
    Else
        AdjustPath = MyPath
    End If
End Function

Private Function GetBase(ByVal FileName As String) As String
    PathRemoveExtension FileName
    GetBase = NullTrim(FileName)
End Function

Public Function GetBasePath(ByVal FullPath As String) As String
    PathRemoveFileSpec FullPath
    GetBasePath = NullTrim(FullPath)
End Function
Public Function GetExt(FileName As String) As String
    GetExt = PtrToStr(PathFindExtension(FileName))
End Function

Private Sub MakeValidFileName(ResName As String)
    Dim n As Integer
    
    For n = 1 To Len(ResName)
        If Not (CBool(PathGetCharType(Asc(Mid(ResName, n, 1))) _
            And GCT_LFNCHAR)) Then
            ResName = Left(ResName, n - 1) & "_" & Mid( _
                ResName, n + 1)
        End If
    Next n
End Sub
Private Function NullTrim(MyString As String) As String
    Dim Pos As Integer
    
    Pos = InStr(MyString, Chr(0))
    
    If Pos Then
        NullTrim = Left(MyString, Pos - 1)
    Else
        NullTrim = MyString
    End If
End Function

Private Function GetTempFile(ResType As String, ByVal ResName As String) As String
    Dim CurFileName As String
    Dim DestFile As String, n As Long
    
    MakeValidFileName ResName
    Do
        n = n + 1
        If ResType <> CStr(RT_HTML) Then
            CurFileName = PtrToStr(PathFindFileName(CurFile)) & IIf( _
                n <> 1, " (" & n & ")", Empty)
        Else
            CurFileName = PtrToStr(PathFindFileName(CurFile)) & " - " & _
                GetBase(ResName) & IIf(n <> 1, " (" & n & ")", Empty) & GetExt(ResName)
        End If
        
        Select Case ResType
            Case RT_BITMAP
                DestFile = DestPath & "Bitmaps\" & CurFileName & ".bmp"
            Case RT_GROUP_ICON
                DestFile = DestPath & "Icons\" & CurFileName & ".ico"
            Case RT_GROUP_CURSOR
                DestFile = DestPath & "Cursors\" & CurFileName & ".cur"
            Case RT_ANICURSOR
                DestFile = DestPath & "Animated Cursors\" & CurFileName & ".ani"
            Case "AVI"
                DestFile = DestPath & "Videos\" & CurFileName & ".avi"
            Case RT_HTML
                DestFile = DestPath & "Web pages\" & CurFileName
        End Select

        If PathFileExists(DestFile) = 0 Then
            Exit Do
        End If
    Loop
    
    GetTempFile = DestFile
End Function
Private Function IsBounded(Arr As Variant) As Boolean
    On Error Resume Next
    IsBounded = IsNumeric(UBound(Arr))
End Function

Private Function PtrToStr(StrPtr As Long) As String
    Dim StrPtrLen As Long

    StrPtrLen = lstrlenPtr(StrPtr)
    PtrToStr = Space(StrPtrLen)

    CopyMemoryToStr PtrToStr, ByVal StrPtr, StrPtrLen
End Function
Public Function EnumResTypeProc(ByVal hModule As Long, ByVal ResType As Long, ByVal lParam As Long) As Long
    Dim EnumNames As Boolean, ResTypeStr As String
    
    If (ResType And &HFFFF0000) = 0 Then
        Select Case ResType
            Case RT_BITMAP
                If ResTypes.Bitmap Then EnumNames = True
            Case RT_GROUP_ICON
                If ResTypes.Icon Then EnumNames = True
            Case RT_GROUP_CURSOR
                If ResTypes.Cursor Then EnumNames = True
            Case RT_ANICURSOR
                If ResTypes.AniCursor Then EnumNames = True
            Case RT_HTML
                If ResTypes.HTML Then EnumNames = True
        End Select
        If EnumNames Then
            EnumResourceNames hModule, ByVal ResType, AddressOf EnumResNameProc, 0
        End If
    Else
        ResTypeStr = PtrToStr(ResType)
        If (ResTypeStr = "AVI") And ResTypes.Video Then
            EnumResourceNames hModule, ByVal ResTypeStr, AddressOf EnumResNameProc, 0
        End If
    End If
    
    EnumResTypeProc = (CancelResSearch = 0)
End Function
Public Function EnumResNameProc(ByVal hModule As Long, ByVal ResType As Long, ByVal ResID As Long, ByVal lParam As Long) As Long
    Dim MyResName As Variant, MyResType As Variant
    Dim ResTypeFinal As RESOURCETYPES
    
    If (ResType And &HFFFF0000) = 0 Then
        MyResType = ResType
    Else
        MyResType = PtrToStr(ResType)
    End If

    If (ResID And &HFFFF0000) = 0 Then
        MyResName = ResID
    Else
        MyResName = PtrToStr(ResID)
    End If
    
    Select Case MyResType
        Case RT_BITMAP:       ResTypeFinal = rcrtBitmap
        Case RT_GROUP_ICON:   ResTypeFinal = rcrtIcon
        Case RT_GROUP_CURSOR: ResTypeFinal = rcrtCursor
        Case RT_ANICURSOR:    ResTypeFinal = rcrtAniCursor
        Case "AVI":           ResTypeFinal = rcrtVideo
        Case RT_HTML:         ResTypeFinal = rcrtHTML
    End Select
    If SaveResource(hModule, CStr(MyResType), CStr(MyResName)) Then
        TriggerObj.ResFound ResTypeFinal, MyResName
    End If
    
    EnumResNameProc = (CancelResSearch = 0)
End Function
Private Function GetResourceData(hModule As Long, ResType As String, ByVal ResName As String) As Byte()
    Dim hResource As Long, hMemGlobal As Long
    Dim ResSize As Long, Data() As Byte
    Dim MemPtr As Long

    If IsNumeric(ResName) Then
        ResName = "#" & ResName
    End If
    If IsNumeric(ResType) Then
        hResource = FindResource(hModule, ResName, ByVal CLng(ResType))
    Else
        hResource = FindResource(hModule, ResName, ByVal ResType)
    End If
    If hResource <> 0 Then
        hMemGlobal = LoadResource(hModule, hResource)
        If hMemGlobal <> 0 Then
            MemPtr = LockResource(hMemGlobal)
            If MemPtr <> 0 Then
                ResSize = SizeofResource(hModule, hResource)
                If ResSize > 0 Then
                    ReDim Data(0 To ResSize - 1)
                    CopyMemory Data(0), ByVal MemPtr, ResSize
                    GetResourceData = Data
                End If
            End If
        End If
        FreeResource hResource
    End If
End Function
Private Sub SaveFileInfo(FileName As String, ResType As String, ResName As String)
    Dim FileNum As Integer, OutFile As String, FileOnly As String

    OutFile = GetBasePath(FileName) & "\" & IdxFile
    FileOnly = PtrToStr(PathFindFileName(FileName))
    
    FileNum = FreeFile
    
    Open OutFile For Append Access Write As #FileNum
    
    Print #FileNum, FileOnly & " - SOURCE FILE = " & CurFile & _
        " RESOURCE TYPE = " & ResType & " RESOURCE NAME = " & ResName
    
    Close #FileNum
End Sub
Private Function SaveResource(hModule As Long, ResType As String, ResName As String) As Boolean
    Dim FileName As String, ProcRet As Boolean, errnumber As Long

    FileName = GetTempFile(ResType, ResName)

    Do
        CancelResSearch = 0
        Select Case ResType
            Case RT_BITMAP
                ProcRet = SaveResourceBitmap(hModule, FileName, ResName)
            Case RT_GROUP_CURSOR
                ProcRet = SaveResourceCursor(hModule, FileName, ResName)
            Case RT_GROUP_ICON
                ProcRet = SaveResourceIcon(hModule, FileName, ResName)
            Case Else
                ProcRet = SaveResourceData(hModule, FileName, ResType, ResName)
        End Select

        Select Case CancelResSearch
            Case 0
                SaveFileInfo FileName, ResType, ResName
                Exit Do
            Case 61
                If PathFileExists(FileName) <> 0 Then
                    Kill FileName
                End If
                If MsgBox("The drive is full." & vbCrLf & _
                    "Please free some space on the drive.", _
                    vbCritical Or vbRetryCancel) = vbCancel Then
                    Exit Do
                End If
            Case Else
                If Len(Trim(Err.Description)) <> 0 Then
                    MsgBox Err.Description, vbCritical
                End If
                Exit Do
        End Select
    Loop
    
    SaveResource = ProcRet
End Function
Private Function SaveResourceData(hModule As Long, FileName As String, ResType As String, ResName As String) As Boolean
    Dim MyData() As Byte, FileNum As Integer
    
    MyData = GetResourceData(hModule, ResType, ResName)
    If Not IsBounded(MyData) Then Exit Function
    
    On Error GoTo File_Err
    FileNum = FreeFile
    
    Open FileName For Binary Lock Write As #FileNum
        Put #FileNum, , MyData
    Close #FileNum
    SaveResourceData = True
    Exit Function
File_Err:
    Close #FileNum
    CancelResSearch = Err.Number
End Function
Private Function SaveResourceBitmap(hModule As Long, FileName As String, ResName As String) As Boolean
    Dim MyData() As Byte, FileNum As Integer
    Dim FileHdr As BITMAPFILEHEADER
    Dim BmpHdr As BITMAPINFOHEADER, TableSize As Long
    
    MyData = GetResourceData(hModule, RT_BITMAP, ResName)
    If Not IsBounded(MyData) Then Exit Function
    
    CopyMemory BmpHdr.biSize, MyData(0), Len(BmpHdr)
    
    With BmpHdr
        If .biSize <> Len(BmpHdr) Then Exit Function
        If .biWidth <= 0 Then Exit Function
        If .biHeight <= 0 Then Exit Function
        If .biPlanes <> 1 Then Exit Function
        Select Case .biBitCount
            Case 1, 4, 8, 16, 24, 32
            Case Else: Exit Function
        End Select
    End With
    
    If BmpHdr.biBitCount <= 8 Then
        If BmpHdr.biClrUsed <> 0 Then
            TableSize = BmpHdr.biClrUsed
        Else
            TableSize = 2 ^ BmpHdr.biBitCount
        End If
    End If
    
    ReDim BmpData(0 To UBound(MyData) + 14)
    With FileHdr
        .bfType = &H4D42
        .bfSize = UBound(BmpData) + 1
        .bfOffBits = Len(FileHdr) + Len(BmpHdr) + TableSize * 4
    End With
    
    On Error GoTo File_Err
    FileNum = FreeFile
    
    Open FileName For Binary Lock Write As #FileNum
        Put #FileNum, , FileHdr
        Put #FileNum, , MyData
    Close #FileNum
    SaveResourceBitmap = True
    Exit Function
File_Err:
    Close #FileNum
    CancelResSearch = Err.Number
End Function
Private Function SaveResourceIcon(hModule As Long, FileName As String, ResName As String) As Boolean
    Dim FileNum As Integer, n As Integer
    Dim GroupData() As Byte, IconCount As Integer
    Dim GrpIconEntries() As GRPICONDIRENTRY
    Dim CurEntry As ICONDIRENTRY
    Dim Offset As Long, GrpHdr(0 To 5) As Byte
    
    On Error GoTo File_Err
    FileNum = FreeFile

    Open FileName For Binary Lock Write As #FileNum
        'Write Dir
        GroupData = GetResourceData(hModule, RT_GROUP_ICON, ResName)
        If Not IsBounded(GroupData) Then Exit Function
        CopyMemory IconCount, GroupData(4), Len(IconCount)
        If GroupData(2) <> 1 Then GoTo Error
        CopyMemory GrpHdr(0), GroupData(0), 6
        Put #FileNum, , GrpHdr
        If IconCount <= 0 Then GoTo Error
        
        'Write Entries
        ReDim GrpIconEntries(0 To IconCount - 1)
        For n = 0 To IconCount - 1
            CopyMemory GrpIconEntries(n), GroupData(6 + n * 14), 14
        Next n
        Offset = 6 + Len(CurEntry) * IconCount
        For n = 0 To IconCount - 1
            With GrpIconEntries(n)
                If .bWidth < 0 Then GoTo Error
                CurEntry.bWidth = .bWidth
                If .bHeight < 0 Then GoTo Error
                CurEntry.bHeight = .bHeight
                If (.bColorCount < 0) Or _
                    (.bColorCount > 16) Then GoTo Error
                CurEntry.bColorCount = .bColorCount
                If .wPlanes <> 1 Then GoTo Error
                CurEntry.wPlanes = .wPlanes
                Select Case .wBitCount
                    Case 1, 4, 8, 16, 24, 32
                    Case Else: GoTo Error
                End Select
                CurEntry.wBitCount = .wBitCount
                If .dwBytesInRes <= 0 Then GoTo Error
                CurEntry.dwBytesInRes = .dwBytesInRes
                CurEntry.dwImageOffset = Offset
                Offset = Offset + .dwBytesInRes
            End With
            Put #FileNum, , CurEntry
        Next n
        
        'Write Image Data
        For n = 0 To IconCount - 1
            Put #FileNum, , GetResourceData(hModule, RT_ICON, _
                GrpIconEntries(n).nID)
        Next n
    Close #FileNum
    SaveResourceIcon = True
    Exit Function
Error:
    Close #FileNum
    Kill FileName
    Exit Function
File_Err:
    Close #FileNum
    CancelResSearch = Err.Number
End Function
Private Function SaveResourceCursor(hModule As Long, FileName As String, ResName As String) As Boolean
    Dim FileNum As Integer, n As Integer
    Dim GroupData() As Byte, CursorCount As Integer
    Dim GrpCursorEntries() As GRPCURSORDIRENTRY
    Dim CurEntry As CURSORDIRENTRY
    Dim Offset As Long, GrpHdr(0 To 5) As Byte
    Dim CurHsptX As Integer, CurHsptY As Integer
    Dim ImgData() As BYTEARRAY, FinalImgData() As Byte
    
    On Error GoTo File_Err
    FileNum = FreeFile
    Open FileName For Binary Lock Write As #FileNum
        'Write Dir
        GroupData = GetResourceData(hModule, RT_GROUP_CURSOR, ResName)
        If Not IsBounded(GroupData) Then Exit Function
        CopyMemory CursorCount, GroupData(4), Len(CursorCount)
        If GroupData(2) <> 2 Then GoTo Error
        CopyMemory GrpHdr(0), GroupData(0), 6
        Put #FileNum, , GrpHdr
        If CursorCount <= 0 Then GoTo Error
        
        'Write Entries
        ReDim GrpCursorEntries(0 To CursorCount - 1)
        ReDim ImgData(0 To CursorCount - 1)
        For n = 0 To CursorCount - 1
            CopyMemory GrpCursorEntries(n), GroupData(6 + n * 14), 14
        Next n
        Offset = 6 + Len(CurEntry) * CursorCount
        For n = 0 To CursorCount - 1
            With GrpCursorEntries(n)
                ImgData(n).Data = GetResourceData(hModule, RT_CURSOR, _
                    GrpCursorEntries(n).wNameOrdinal)
                If IsBounded(ImgData(n).Data) Then
                    CopyMemory CurHsptX, ImgData(n).Data(0), Len(CurHsptX)
                    CopyMemory CurHsptY, ImgData(n).Data(2), Len(CurHsptY)
                Else
                    GoTo Error
                End If
            
                If (.wWidth <= 0) Or (.wWidth > 255) Then GoTo Error
                CurEntry.bWidth = .wWidth
                If (.wHeight <= 0) Or (.wHeight > 127) Then GoTo Error
                CurEntry.bHeight = .wHeight \ 2
                CurEntry.bColorCount = 2
                If .wPlanes <> 1 Then GoTo Error
                CurEntry.wHotspotX = CurHsptX
                CurEntry.wHotspotY = CurHsptY
                If .lBytesInRes <= 0 Then GoTo Error
                CurEntry.dwBytesInRes = .lBytesInRes
                CurEntry.dwImageOffset = Offset

                Offset = Offset + .lBytesInRes
            End With
            Put #FileNum, , CurEntry
        Next n
        
        'Write Image Data
        For n = 0 To CursorCount - 1
            ReDim FinalImgData(0 To UBound(ImgData(n).Data) - 4)
            CopyMemory FinalImgData(0), ImgData(n).Data(4), _
                UBound(FinalImgData) + 1
            Put #FileNum, , FinalImgData
        Next n
    Close #FileNum
    SaveResourceCursor = True
    Exit Function
Error:
    Close #FileNum
    Kill FileName
    Exit Function
File_Err:
    Close #FileNum
    CancelResSearch = Err.Number
End Function
Public Function SearchForResources(FileName As String, MyResTypes As RESTYPEOPTIONS, Obj As Object, Destination As String) As Boolean
    Dim hModule As Long
    
    CancelResSearch = False
    ResTypes = MyResTypes
    Set TriggerObj = Obj
    DestPath = AdjustPath(Destination)
    CurFile = FileName
    
    hModule = LoadLibraryEx(FileName, 0, LOAD_LIBRARY_AS_DATAFILE)
    If hModule <> 0 Then
        EnumResourceTypes hModule, AddressOf EnumResTypeProc, 0
        FreeLibrary hModule
    End If
    
    SearchForResources = (CancelResSearch <> 0)
End Function
