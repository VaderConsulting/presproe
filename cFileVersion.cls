VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFileVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FileVersionMS As Long
Private FileVersionLS As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type VS_FIXEDFILEINFO
    Signature As Long
    StrucVersion As Long         '  e.g. 0x00000042 = "0.42"
    FileVersionMS As Long        '  e.g. 0x00030075 = "3.75"
    FileVersionLS As Long        '  e.g. 0x00000031 = "0.31"
    ProductVersionMS As Long     '  e.g. 0x00030010 = "3.10"
    ProductVersionLS As Long     '  e.g. 0x00000031 = "0.31"
    FileFlagsMask As Long        '  = 0x3F for version "0.42"
    FileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    FileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    FileType As Long             '  e.g. VFT_DRIVER
    FileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    FileDateMS As Long           '  e.g. 0
    FileDateLS As Long           '  e.g. 0
End Type
Public Property Get FileVersion() As String
    FileVersion = HiWord(FileVersionMS) & "." & _
        LoWord(FileVersionMS) & "." & HiWord( _
        FileVersionLS) & "." & LoWord(FileVersionLS)
End Property

Private Function Get64Bit(DWord1 As Long, DWord2 As Long) As Long
    Dim Buffer As Currency
    
    CopyMemory Buffer, DWord1, Len(DWord1)
    CopyMemory VarPtr(Buffer) + 4, DWord2, Len(DWord2)
    
    On Error Resume Next
    Get64Bit = CLng(Buffer * 10000)
End Function
Public Function AreFilesEqual(File1 As String, File2 As String) As Boolean
    Dim Ver1 As Double, Ver2 As Double
    
    Ver1 = OpenFile(File1)
    If Ver1 = 0 Then Exit Function
    Ver2 = OpenFile(File2)
    If Ver2 = 0 Then Exit Function
    
    If Ver1 = Ver2 Then AreFilesEqual = True
End Function
Public Function OpenFile(FileName As String) As Long
    Dim VersionInfoSize As Long, Dummy As Long
    Dim VerBuf() As Byte
    Dim VerLoc As Long, VerInfo As VS_FIXEDFILEINFO
    
    VersionInfoSize = GetFileVersionInfoSize(FileName, Dummy)
    
    If VersionInfoSize <> 0 Then
        ReDim VerBuf(0 To VersionInfoSize - 1)
        If GetFileVersionInfo(FileName, 0, VersionInfoSize, _
            VerBuf(0)) <> 0 Then
            If VerQueryValue(VerBuf(0), "\", VerLoc, Len(VerInfo)) <> 0 Then
                If VerLoc <> 0 Then
                    CopyMemory VerInfo.Signature, ByVal VerLoc, Len(VerInfo)
                    
                    FileVersionMS = VerInfo.FileVersionMS
                    FileVersionLS = VerInfo.FileVersionLS
                    
                    OpenFile = Get64Bit(FileVersionMS, FileVersionLS)
                End If
            End If
        End If
    End If
End Function
Function LoWord(ByVal DWord As Long) As Integer
    If DWord And &H8000& Then
        LoWord = DWord Or &HFFFF0000
    Else
        LoWord = DWord And &HFFFF&
    End If
End Function
Function HiWord(ByVal DWord As Long) As Integer
    HiWord = (DWord And &HFFFF0000) \ 65536
End Function
Public Property Get VersionMajor() As Long
    VersionMajor = FileVersionMS
End Property


Public Property Get VersionMinor() As Long
    VersionMinor = FileVersionLS
End Property
