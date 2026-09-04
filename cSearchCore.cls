VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSearchCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FileCount As Long
Private Cancel As Boolean
Private DirCount As Long

Public Event FileFound(FileName As String, FileCount As Long, DirCount As Long, Cancel As Boolean)
Public Event SearchFinished()

Private Const MAX_PATH = 260

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
   dwFileAttributes As Long
   ftCreationTime As FILETIME
   ftLastAccessTime As FILETIME
   ftLastWriteTime As FILETIME
   nFileSizeHigh As Long
   nFileSizeLow As Long
   dwReserved0 As Long
   dwReserved1 As Long
   cFileName As String * MAX_PATH
   cAlternate As String * 14
End Type
Private Const INVALID_HANDLE_VALUE = -1
Private Const FindExInfoStandard = 0
Private Const FindExSearchLimitToDirectories = 2

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Const DRIVE_FIXED = 3
Private Declare Sub CopyMemoryToStr Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As String, Source As Any, ByVal Length As Long)
Private Declare Function lstrlenPtr Lib "kernel32" Alias "lstrlenA" (ByVal lpLong As Long) As Long
Private Declare Function PathMatchSpec Lib "shlwapi" Alias "PathMatchSpecA" (ByVal pszFileParam As String, ByVal pszSpec As String) As Long
Private Declare Function PathFindFileName Lib "shlwapi" Alias "PathFindFileNameA" (ByVal pPath As String) As Long
Private Declare Function FindFirstFileEx Lib "kernel32" Alias "FindFirstFileExA" (ByVal lpFileName As String, ByVal fInfoLevelId As Long, lpFindFileData As WIN32_FIND_DATA, ByVal fSearchOp As Long, ByVal lpSearchFilter As Long, ByVal dwAdditionalFlags As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformID As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const CSIDL_BITBUCKET = &HA
Private Const S_OK = 0
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal PIDL As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, PIDL As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Function AcceptFile(FilePath As String, FileExtArray() As String) As Boolean
    Dim n As Integer, FileName As String
    
    FileName = PtrToStr(PathFindFileName(FilePath))
    
    For n = LBound(FileExtArray) To UBound(FileExtArray)
        If PathMatchSpec(FileName, FileExtArray(n)) <> 0 Then
            AcceptFile = True
        End If
    Next n
End Function
Public Function GetDirCount(ByVal AllDrives As Boolean, Optional ByVal BaseDir As String = "/") As Long
    Dim Buffer As String, MyDrives() As String, n As Integer
    Dim DummyArray() As String
    
    If (Not AllDrives) And (BaseDir = "/") Then AllDrives = True
    
    DirCount = 0
    Buffer = Space(MAX_PATH)
    
    If AllDrives Then
        GetLogicalDriveStrings Len(Buffer) - 1, Buffer
    
        MyDrives = Split(Buffer, Chr(0))
        Cancel = False
        
        For n = LBound(MyDrives) To UBound(MyDrives)
            If GetDriveType(MyDrives(n)) = DRIVE_FIXED Then
                IntSearchForFiles MyDrives(n), True, DummyArray, False, True
            End If
        Next n
    Else
        IntSearchForFiles BaseDir, True, DummyArray, False, True
    End If
    
    GetDirCount = DirCount
End Function
Private Function IsNT4OrHigher() As Boolean
    Dim VersionInfo As OSVERSIONINFO
    
    VersionInfo.dwOSVersionInfoSize = Len(VersionInfo)
    
    GetVersionEx VersionInfo
    
    If (VersionInfo.dwPlatformID = VER_PLATFORM_WIN32_NT) And _
        (VersionInfo.dwMajorVersion >= 4) Then
        IsNT4OrHigher = True
    End If
End Function
Private Function PtrToStr(StrPtr As Long) As String
    Dim StrPtrLen As Long

    StrPtrLen = lstrlenPtr(StrPtr)
    PtrToStr = Space(StrPtrLen)

    CopyMemoryToStr PtrToStr, ByVal StrPtr, StrPtrLen
End Function

Private Function ExtensionInMask(Extension As String, ExtArray() As String) As Boolean
    Dim n As Integer
    
    For n = LBound(ExtArray) To UBound(ExtArray)
        If Extension = ExtArray(n) Then
            ExtensionInMask = True
            Exit Function
        End If
    Next n
End Function
Private Function GetExtArray(ByVal FileMask As String) As String()
    Dim MyArray() As String
    
    FileMask = Trim(FileMask)
    If Len(FileMask) = 0 Then Exit Function
    
    If InStr(FileMask, "; ") Then
        FileMask = Replace(FileMask, "; ", "|")
    ElseIf InStr(FileMask, ";") Then
        FileMask = Replace(FileMask, ";", "|")
    ElseIf InStr(FileMask, ", ") Then
        FileMask = Replace(FileMask, ", ", "|")
    ElseIf InStr(FileMask, ",") Then
        FileMask = Replace(FileMask, ",", "|")
    ElseIf InStr(FileMask, " ") Then
        FileMask = Replace(FileMask, " ", "|")
    End If
    
    If InStr(FileMask, "|") Then
        MyArray = Split(FileMask, "|")
    Else
        ReDim MyArray(0 To 0)
        MyArray(0) = FileMask
    End If
    
    GetExtArray = MyArray
End Function
Private Function IntSearchForFiles95(FileRoot As String, Recurse As Boolean, FileExtArray() As String, ResetCounter As Boolean, GetFolders As Boolean) As Boolean
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    Dim MyPath As String, Root As String, Tmp As String
    
    Root = AdjustPath(FileRoot)
    MyPath = Root & "*.*"
    If ResetCounter Then FileCount = 0: DirCount = 0
    
    hFile = FindFirstFile(MyPath, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then
        If Not GetFolders Then
            Cancel = ListFile(FileRoot, FileExtArray)
        End If
        Do
            If Cancel Then Exit Do
            If (WFD.dwFileAttributes And vbDirectory) And _
                (Not (WFD.dwFileAttributes And vbHidden)) And _
                (Not (WFD.dwFileAttributes And vbSystem)) Then
                DirCount = DirCount + 1
                If Recurse Then
                    Tmp = NullTrim(WFD.cFileName)
                    If (Tmp <> ".") And (Tmp <> "..") Then
                        FileRoot = Root & Tmp
                        Cancel = IntSearchForFiles95(FileRoot, Recurse, _
                            FileExtArray, False, GetFolders)
                    End If
                End If
            End If
            If Cancel Then Exit Do
        Loop While FindNextFile(hFile, WFD)
        FindClose hFile
    End If
    
    IntSearchForFiles95 = Cancel
End Function
Private Function IntSearchForFiles(FileRoot As String, Recurse As Boolean, FileExtArray() As String, ResetCounter As Boolean, GetFolders As Boolean) As Boolean
    If IsNT4OrHigher Then
        IntSearchForFiles = IntSearchForFilesNT(FileRoot, Recurse, _
            FileExtArray, ResetCounter, GetFolders)
    Else
        IntSearchForFiles = IntSearchForFiles95(FileRoot, Recurse, _
            FileExtArray, ResetCounter, GetFolders)
    End If
End Function
Private Function IntSearchForFilesNT(FileRoot As String, Recurse As Boolean, FileExtArray() As String, ResetCounter As Boolean, GetFolders As Boolean) As Boolean
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    Dim MyPath As String, Root As String, Tmp As String
    
    Root = AdjustPath(FileRoot)
    MyPath = Root & "*.*"
    If ResetCounter Then FileCount = 0: DirCount = 0
    
    hFile = FindFirstFileEx(MyPath, FindExInfoStandard, WFD, _
        FindExSearchLimitToDirectories, 0, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        If Not GetFolders Then
            Cancel = ListFile(FileRoot, FileExtArray)
        End If
        Do
            If Cancel Then Exit Do
            If (WFD.dwFileAttributes And vbDirectory) And _
                (Not (WFD.dwFileAttributes And vbHidden)) And _
                (Not (WFD.dwFileAttributes And vbSystem)) Then
                DirCount = DirCount + 1
                If Recurse Then
                    Tmp = NullTrim(WFD.cFileName)
                    If (Tmp <> ".") And (Tmp <> "..") Then
                        FileRoot = Root & Tmp
                        Cancel = IntSearchForFilesNT(FileRoot, Recurse, _
                            FileExtArray, False, GetFolders)
                    End If
                End If
            End If
            If Cancel Then Exit Do
        Loop While FindNextFile(hFile, WFD)
        FindClose hFile
        IntSearchForFilesNT = Cancel
    Else
        IntSearchForFilesNT = IntSearchForFiles95(FileRoot, Recurse, _
            FileExtArray, ResetCounter, GetFolders)
    End If
End Function
Public Sub SearchForFiles(ByVal FileRoot As String, Recurse As Boolean, FileMask As String)
    Dim FileExtArray() As String
    
    FileExtArray = GetExtArray(FileMask)
    Cancel = False
    IntSearchForFiles FileRoot, Recurse, FileExtArray, True, False
    RaiseEvent SearchFinished
End Sub


Public Sub SearchSystemForFiles(FileMask As String)
    Dim Buffer As String, MyDrives() As String, n As Integer
    Dim FileExtArray() As String
    
    FileExtArray = GetExtArray(FileMask)
    Buffer = Space(MAX_PATH)
    
    GetLogicalDriveStrings Len(Buffer) - 1, Buffer

    MyDrives = Split(Buffer, Chr(0))
    Cancel = False
    
    DirCount = 0: FileCount = 0
    
    For n = LBound(MyDrives) To UBound(MyDrives)
        If GetDriveType(MyDrives(n)) = DRIVE_FIXED Then
            If IntSearchForFiles(MyDrives(n), True, FileExtArray, False, False) Then
                Exit For
            End If
        End If
    Next n
    
    RaiseEvent SearchFinished
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
Private Function ListFile(FileRoot As String, FileExtArray() As String) As Boolean
    Dim WFD As WIN32_FIND_DATA, hFile As Long
    Dim MyPath As String, Root As String, Tmp As String
    Dim MyFile As String
    
    Root = AdjustPath(FileRoot)
    MyPath = Root & "*.*"
    
    hFile = FindFirstFile(MyPath, WFD)
    If hFile <> INVALID_HANDLE_VALUE Then
        Do
            Tmp = NullTrim(WFD.cFileName)

            If Not (WFD.dwFileAttributes And vbDirectory) _
                = vbDirectory Then
                MyFile = Root & Tmp
                If AcceptFile(MyFile, FileExtArray) Then
                    FileCount = FileCount + 1
                    RaiseEvent FileFound(MyFile, FileCount, DirCount, Cancel)
                    If Cancel Then Exit Do
                End If
            End If
        Loop While FindNextFile(hFile, WFD)

        hFile = FindClose(hFile)
    End If
    
    ListFile = Cancel
End Function
Private Function AdjustPath(MyPath As String) As String
    If Right(MyPath, 1) <> "\" Then
        AdjustPath = MyPath & "\"
    Else
        AdjustPath = MyPath
    End If
End Function

