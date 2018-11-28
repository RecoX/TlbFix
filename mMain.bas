Attribute VB_Name = "mMain"

Option Explicit

Private Const NO_ERROR          As Long = 0
Private Const MAX_PATH          As Long = 260
Private Const CSIDL_SYSTEM      As Long = &H25
Private Const CSIDL_SYSTEMX86   As Long = &H29

Private Type SHITEMID
    cb      As Long
    abID    As Byte
End Type

Private Type ITEMIDLIST
    mkid    As SHITEMID
End Type

Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

Public Sub Main()
    Dim sFile           As String
    Dim objTypeLibInfo  As TypeLibInfo
    Dim lRet            As Long
    
    On Error GoTo Main_Error

    Call IsWow64Process(GetCurrentProcess, lRet)
    
    If lRet = 0 Then
        sFile = GetSpecialfolder(CSIDL_SYSTEM) & "msdatsrc.tlb"
    Else
        sFile = GetSpecialfolder(CSIDL_SYSTEMX86) & "msdatsrc.tlb"
    End If
    
    If Not PathFileExists(sFile) = 0 Then
        Set objTypeLibInfo = TLIApplication.TypeLibInfoFromFile(sFile)
        Call objTypeLibInfo.Register

        Call MsgBox("Type library registered successfully. ", vbInformation, "")
    Else
        Call MsgBox("File 'msdatsrc.tlb' not found", vbExclamation, "")
    End If
    
    On Error GoTo 0
    Exit Sub
Main_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Main of Módulo mMain"
End Sub

Private Function GetSpecialfolder(CSIDL As Long) As String
    Dim tITEMIDLIST As ITEMIDLIST
    Dim sPath       As String

    If SHGetSpecialFolderLocation(0, CSIDL, tITEMIDLIST) = NO_ERROR Then

        sPath = Space$(MAX_PATH)

        If SHGetPathFromIDList(ByVal tITEMIDLIST.mkid.cb, ByVal sPath) Then
            GetSpecialfolder = AddBackslash(sPath)
        End If
    End If
End Function

Private Function AddBackslash(ByVal sPath As String) As String
    sPath = sPath & Space(1)
    Call PathAddBackslash(sPath)
    AddBackslash = Left$(sPath, lstrlen(sPath))
End Function
