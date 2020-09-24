Attribute VB_Name = "modRoutines"
Option Explicit

Public Enum WalkDirConsts
  dwWalkDirNormal = 0
  dwWalkCurDirOnly = 1
  dwDirectoriesOnly = 2
End Enum
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
Public Const SW_MAXIMIZE = 3
Public CancelDirSearch As Boolean
Public DriveLetter As String
Public FileLoc As String

Public Sub SearchFromDir(sSrcDir As String, sSearchItem As String, C As ListView, Optional fFlags As WalkDirConsts = dwWalkDirNormal)
  On Error GoTo err

  Dim SubDirCount  As Integer
  Dim TempFilename As String
  Dim Count        As Integer
  Dim Item         As ListItem
  Dim FileSize     As Long
  Dim SubDirs()    As String

  CancelDirSearch = False
    
  If fFlags And dwDirectoriesOnly Then
    'Look for directory entries Only
    TempFilename = Dir$(sSrcDir & sSearchItem, vbDirectory)
    Do While (Len(TempFilename) > 0) And (Not CancelDirSearch)
      'Ignore the "." and ".." directory entries
      If Left$(TempFilename, 1) <> "." Then
        If IsDirectory(sSrcDir & TempFilename) Then
          Set Item = C.ListItems.Add(, , C.ListItems.Count + 1)
          Item.SubItems(1) = sSrcDir
          FileSize = FileLen(sSrcDir & "/" & TempFilename)
          Item.SubItems(1) = TempFilename
          If FileSize >= 0 And FileSize < 1024 Then Item.SubItems(3) = CInt(FileSize) & " B"
          If FileSize >= 1024 And FileSize < 1048576 Then Item.SubItems(3) = Format(FileSize / 1024, "0") & " K"
          If FileSize >= 1048576 And FileSize < 1073741824 Then Item.SubItems(3) = Format(FileSize / 1048576, "0.00") & " MB"
          If FileSize >= 1073741824 Then Item.SubItems(3) = Format(FileSize / 1048576, "0.00") & " GB"
          Item.SubItems(4) = Format(FileDateTime(sSrcDir & "/" & TempFilename), "dd/mm/yyyy")
        End If
      End If
      TempFilename = Dir$
      DoEvents
    Loop
  
  Else
    'Look for non-directory entries
     TempFilename = Dir$(sSrcDir$ & sSearchItem)
    Do While (Len(TempFilename) > 0) And (Not CancelDirSearch)
      Set Item = C.ListItems.Add(, , C.ListItems.Count + 1)
      Item.SubItems(1) = TempFilename
      Item.SubItems(2) = sSrcDir
      FileSize = FileLen(sSrcDir & TempFilename)
      If FileSize >= 0 And FileSize < 1024 Then Item.SubItems(3) = CInt(FileSize) & " B"
      If FileSize >= 1024 And FileSize < 1048576 Then Item.SubItems(3) = Format(FileSize / 1024, "0") & " K"
      If FileSize >= 1048576 And FileSize < 1073741824 Then Item.SubItems(3) = Format(FileSize / 1048576, "0.00") & " MB"
      If FileSize >= 1073741824 Then Item.SubItems(3) = Format(FileSize / 1048576, "0.00") & " GB"
      Item.SubItems(4) = Format(FileDateTime(sSrcDir & TempFilename), "dd/mm/yyyy")
      TempFilename = Dir$
      DoEvents
    Loop
  End If
  
  If (fFlags And dwWalkCurDirOnly) Or CancelDirSearch Then Exit Sub
  
  'Now look for sub-directories
  ReDim SubDirs(10)
  TempFilename = Dir$(sSrcDir$ + "*.*", vbDirectory)
  Do While Len(TempFilename) > 0
    If Left$(TempFilename, 1) <> "." Then
      If IsDirectory(sSrcDir & TempFilename) Then
        SubDirs(SubDirCount) = TempFilename
        SubDirCount = SubDirCount + 1
        If SubDirCount = UBound(SubDirs) Then
          ReDim Preserve SubDirs(SubDirCount + 10)
        End If
      End If
    End If
    TempFilename = Dir$
    DoEvents
    If CancelDirSearch Then Exit Sub
  Loop
  
  'Now walk the subdirectories:
  For Count = 0 To SubDirCount% - 1
    Call SearchFromDir(sSrcDir + SubDirs(Count) + "\", sSearchItem, C, fFlags)
    If CancelDirSearch Then Exit Sub
  Next Count
  FileLoc = sSrcDir
  Exit Sub
err:
    If err.Number = 52 Then
        MsgBox "Nuk ka disk në paisjen përkatëse", vbInformation, "Futeni Diskun!"
    End If
End Sub

Public Function IsDirectory(sFile As String) As Boolean
  Dim nAttr As Integer
  Dim nErr  As Integer

  On Error Resume Next
  nAttr = GetAttr(sFile)
  nErr = err.Number
  
  On Error GoTo 0
  IsDirectory = (nErr = 0) And ((nAttr And vbDirectory) = vbDirectory)
End Function

