Attribute VB_Name = "mCollection"
Option Explicit

'''''''''''''''''''''''''''''''''
'' 22.2.2000 Boris Gillitzer
'''''''''''''''''''''''''''''''''
Private Store As String
 
Public Sub CollectionAdd(ID As String, Value As String)
    Store = Store & UCase(ID) & "=" & Value & "|"
    Debug.Print Store
End Sub
Public Sub CollectionErase()
    Store = Empty
End Sub

Public Function CollectionRemove(ID As String) As Boolean
    Dim Pos As Long, Pos2 As Long, s As String
    
    Pos = InStr(1, Store, ID)
    
    If Pos = 0 Then
        CollectionRemove = False
        Exit Function
    End If
    
    Pos2 = InStr(Pos, Store, "|")
    s = Mid(Store, Pos, Pos2)
    Store = Replace(Store, s, "")
End Function
Public Function CollectionCheck(ByVal ID As String) As String
    Dim Pos As Long, Pos2 As Long, s As String
    Dim Check As String
    
    ID = UCase(ID)
    Pos = InStr(1, Store, ID)
    
    If Pos = 0 Then Exit Function
    
    Pos2 = InStr(Pos, Store, "|")
    s = Mid(Store, Pos, Pos2)
    Pos = InStr(1, s, "=")
    Check = Mid(s, Pos + 1, Len(s) - Pos - 1)
    CollectionCheck = Check
End Function
