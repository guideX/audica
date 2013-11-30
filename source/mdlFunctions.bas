Attribute VB_Name = "mdlFunctions"
Option Explicit

Public Function ReadFile(lFile As String) As String
On Local Error Resume Next

Dim o As Integer, msg As String
o = FreeFile
Open lFile For Input As #o
    'ReadFile = Input(LOF(o), o)
    msg = StrConv(InputB(LOF(o), o), vbUnicode)
    ReadFile = Left(msg, Len(msg) - 2)
Close #o
End Function

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error Resume Next

Dim msg As String
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
End Function

Public Function ReturnDirectoryPath(lFiletype As eFiletypes)
On Local Error Resume Next
Dim i As Integer

If lFiletype <> 0 Then
    For i = 1 To 6
        If lSettings.sDirectorys(i).dFiletype = lFiletype Then
            ReturnDirectoryPath = i
        End If
    Next i
End If
End Function

Public Function ParseString(lWhole As String, lStart As String, lEnd As String)
'On Local Error GoTo ErrHandler

Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
len1 = InStr(lWhole, lStart)
len2 = InStr(lWhole, lEnd)
Str1 = Right(lWhole, Len(lWhole) - len1)
Str2 = Right(lWhole, Len(lWhole) - len2)
ParseString = Left(Str1, Len(Str1) - Len(Str2) - 1)

ErrHandler:
End Function

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
If lInterface.iOS = Windows_98 Or lInterface.iOS = Windows_95 Then Exit Function
Dim msg As Long
'On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  msg = msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If Err Then
  MakeTransparent = 2
End If
End Function

Public Function MakeOpaque(ByVal hwnd As Long) As Long
If lInterface.iOS = Windows_98 Or lInterface.iOS = Windows_95 Then Exit Function
Dim msg As Long
On Error Resume Next
msg = GetWindowLong(hwnd, GWL_EXSTYLE)
msg = msg And Not WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, msg
SetLayeredWindowAttributes hwnd, 0, 0, LWA_ALPHA
MakeOpaque = 0
If Err Then
  MakeOpaque = 2
End If
End Function
