Attribute VB_Name = "More"
Option Explicit
DefInt A-Z
'Get the disk free space
Public Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
'Get the volume information
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
'Get the logical drive strings
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Get the drive type
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'FindFiles functions
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Const MAX_PATH = 260, MAXDWORD = &HFFFF, INVALID_HANDLE_VALUE = -1, FILE_ATTRIBUTE_ARCHIVE = &H20, FILE_ATTRIBUTE_DIRECTORY = &H10, FILE_ATTRIBUTE_HIDDEN = &H2, FILE_ATTRIBUTE_NORMAL = &H80, FILE_ATTRIBUTE_READONLY = &H1, FILE_ATTRIBUTE_SYSTEM = &H4, FILE_ATTRIBUTE_TEMPORARY = &H100
Public Declare Function SendMessageList Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

Public m_cDiskSize As Currency
Public m_cDiskUsed As Currency
Public m_cDiskFree As Currency
Public m_fFreePercent As Single
Public m_lSerial As Long
Public m_sVolume As String
Public m_sFileSystem As String
Public m_sAllDrives As String
Public m_sDriveType As String
Public Const FS_CASE_IS_PRESERVED = &H2
Public Const FS_CASE_SENSITIVE = &H1
Public Const FS_UNICODE_STORED_ON_DISK = &H4
Public Const FS_PERSISTENT_ACLS = &H8
Public Const FS_FILE_COMPRESSION = &H10
Public Const FS_VOL_IS_COMPRESSED = &H8000

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Public Type WIN32_FIND_DATA
    dwFileAttributes As Long:    ftCreationTime As FILETIME:    ftLastAccessTime As FILETIME:    ftLastWriteTime As FILETIME:    nFileSizeHigh As Long:    nFileSizeLow As Long:    dwReserved0 As Long:    dwReserved1 As Long:    cFileName As String * MAX_PATH:    cAlternate As String * 14
End Type

Const LB_ADDSTRING = &H180

Public Function FindFilesInDirectory(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, Optional handleListbox As Long, Optional SendBackInFilename_PathFormat As Boolean, Optional DontLookInSubFolders As Boolean)
Dim Filename As String
Dim DirName As String
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim temp As String
Dim dirNames() As String
Dim nDir As Integer
If Right(path, 1) <> "\" Then path = path & "\"
If Not DontLookInSubFolders Then
nDir = 0
ReDim dirNames(nDir)
Cont = True
hSearch = FindFirstFile(path & "*", WFD)
If hSearch <> INVALID_HANDLE_VALUE Then
    Do While Cont
    DirName = StripNulls(WFD.cFileName)
    If (DirName <> ".") And (DirName <> "..") Then
        If GetFileAttributes(path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
            dirNames(nDir) = DirName
            DirCount = DirCount + 1
            nDir = nDir + 1
            ReDim Preserve dirNames(nDir)
        End If
    End If
    Cont = FindNextFile(hSearch, WFD)
    DoEvents
    Loop
    Cont = FindClose(hSearch)
End If
End If
hSearch = FindFirstFile(path & SearchStr, WFD)
Cont = True
If hSearch <> INVALID_HANDLE_VALUE Then
    While Cont
    Filename = StripNulls(WFD.cFileName)
    DoEvents
    If (Filename <> ".") And (Filename <> "..") Then
        FindFilesInDirectory = FindFilesInDirectory + (WFD.nFileSizeHigh * MAXDWORD)
        FileCount = FileCount + 1
        temp = path & Filename
 If SendBackInFilename_PathFormat Then temp = Filename & vbTab & path
 SendMessageList handleListbox, LB_ADDSTRING, -1, ByVal temp
    End If
    Cont = FindNextFile(hSearch, WFD)
    Wend
    Cont = FindClose(hSearch)
End If
If Not DontLookInSubFolders Then
If nDir > 0 Then
    For i = 0 To nDir - 1
    DoEvents
    Debug.Print path & dirNames(i) & "\"
    If Not SendBackInFilename_PathFormat Then
    FindFilesInDirectory = FindFilesInDirectory + FindFilesInDirectory(path & dirNames(i) & "\", SearchStr, FileCount, DirCount, handleListbox)
    Else
    FindFilesInDirectory = FindFilesInDirectory + FindFilesInDirectory(path & dirNames(i) & "\", SearchStr, FileCount, DirCount, handleListbox, True)
    End If
    Next i
End If
End If
End Function

Private Function StripNulls(OriginalStr As String) As String
If (InStr(OriginalStr, Chr(0)) > 0) Then
    OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
End If
StripNulls = OriginalStr
End Function

Public Property Get cDiskSize() As Currency
    cDiskSize = m_cDiskSize
End Property

Public Property Get cDiskUsed() As Currency
    cDiskUsed = m_cDiskUsed
End Property

Public Property Get cDiskFree() As Currency
    cDiskFree = m_cDiskFree
End Property

Public Property Get fFreePercent() As Single
    fFreePercent = m_fFreePercent
End Property

Public Sub GetVolumeInfo(ByVal sDrive As String)
    Dim sBuffer As String
    Dim sSysName As String
    Dim lResult As Long
    Dim lSysFlags As Long
    Dim lComponentLength As Long
    
    sBuffer = String$(256, 0)
    sSysName = String$(256, 0)
    lResult = GetVolumeInformation(sDrive, sBuffer, 255, m_lSerial, lComponentLength, lSysFlags, sSysName, 255)
    
    If lResult = 0 Then
        m_sVolume = "Unable To retrieve information"
        m_sFileSystem = "Unable To retrieve information"
        m_lSerial = 0
    Else
        m_sVolume = Left$(sBuffer, InStr(sBuffer, Chr$(0)) - 1)
        m_sFileSystem = Left$(sSysName, InStr(sSysName, Chr$(0)) - 1)
    End If
End Sub

Public Property Get lSerial() As Long
    lSerial = m_lSerial
End Property

Public Property Get sAllDrives() As String
    sAllDrives = m_sAllDrives
End Property

Public Property Get sDriveType() As String
    sDriveType = m_sDriveType
End Property

Public Property Get sSerial() As String
    sSerial = Hex$(m_lSerial)
End Property

Public Property Get sVolume() As String
    sVolume = m_sVolume
End Property

Public Property Get sFileSystem() As String
    sFileSystem = m_sFileSystem
End Property

Public Sub Class_Initialize()
    Dim sTemp As String
    Dim iPos As Integer
    
    sTemp = String$(2048, 0)
    Call GetLogicalDriveStrings(2047, sTemp)
    
    m_sAllDrives = ""

    Do
        iPos = InStr(sTemp, Chr$(0))

        If iPos > 1 Then

            If m_sAllDrives = "" Then
                m_sAllDrives = Left$(sTemp, iPos - 1)
            Else
                m_sAllDrives = m_sAllDrives & "," & Left$(sTemp, iPos - 1)
            End If
            sTemp = Mid$(sTemp, iPos + 1)
        End If
    Loop Until iPos <= 1
End Sub



