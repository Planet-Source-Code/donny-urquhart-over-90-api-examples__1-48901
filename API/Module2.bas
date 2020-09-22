Attribute VB_Name = "Registry"
Option Explicit
Option Compare Text

'close an open registry key
Private Declare Function RegCloseKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long
             
'connect with the registry on a remote machine
Private Declare Function RegConnectRegistry _
        Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
            (ByVal lpMachineName As String, _
             ByVal hKey As Long, _
             phkResult As Long) _
             As Long

'create a new registry key
Private Declare Function RegCreateKey _
        Lib "advapi32.dll" _
        Alias "RegCreateKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             phkResult As Long) _
             As Long
'create new - entended
Private Declare Function RegCreateKeyEx _
        Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal Reserved As Long, _
             ByVal lpClass As String, _
             ByVal dwOptions As Long, _
             ByVal samDesired As Long, _
             lpSecurityAttributes As SECURITY_ATTRIBUTES, _
             phkResult As Long, _
             lpdwDisposition As Long) _
             As Long

'delete the specified registry key (also any sub keys
'for non-NT based systems)
Private Declare Function RegDeleteKey _
        Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String) _
             As Long

'delete a registry value
Private Declare Function RegDeleteValue _
        Lib "advapi32.dll" _
        Alias "RegDeleteValueA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String) _
             As Long

'return a list of registry sub keys in the specified key
Private Declare Function RegEnumKey _
        Lib "advapi32.dll" _
        Alias "RegEnumKeyA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpName As String, _
             ByVal cbName As Long) _
             As Long
Private Declare Function RegEnumKeyEx _
        Lib "advapi32.dll" _
        Alias "RegEnumKeyExA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpName As String, _
             lpcbName As Long, _
             ByVal lpReserved As Long, _
             ByVal lpClass As String, _
             lpcbClass As Long, _
             lpftLastWriteTime As FILETIME) _
             As Long

'get a list of registry values in a key
Private Declare Function RegEnumValue _
        Lib "advapi32.dll" _
        Alias "RegEnumValueA" _
            (ByVal hKey As Long, _
             ByVal dwIndex As Long, _
             ByVal lpValueName As String, _
             lpcbValueName As Long, _
             ByVal lpReserved As Long, _
             lpType As Long, _
             lpData As Byte, _
             lpcbData As Long) _
             As Long

'writes all the attributes of the specified open key
'into the registry
Private Declare Function RegFlushKey _
        Lib "advapi32.dll" _
            (ByVal hKey As Long) _
             As Long

'get the security attributes of the specified key
Private Declare Function RegGetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR, _
             lpcbSecurityDescriptor As Long) _
             As Long

'creates a subkey under HKEY_USER or HKEY_LOCAL_MACHINE
'and stores registration information from a specified
'file into that subkey. This registration information
'is in the form of a hive. A hive is a discrete body of
'keys, subkeys, and values that is rooted at the top of
'the registry hierarchy. A hive is backed by a single
'file and .LOG file
Private Declare Function RegLoadKey _
        Lib "advapi32.dll" _
        Alias "RegLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpFile As String) _
             As Long

'notify a specified procedure (use the AddressOf
'operator), that a key has changed
Private Declare Function RegNotifyChangeKeyValue _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal bWatchSubtree As Long, _
             ByVal dwNotifyFilter As Long, _
             ByVal hEvent As Long, _
             ByVal fAsynchronus As Long) _
             As Long

'open a registry key for access
Private Declare Function RegOpenKey _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             phkResult As Long) _
             As Long
Private Declare Function RegOpenKeyEx _
        Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal ulOptions As Long, _
             ByVal samDesired As Long, _
             phkResult As Long) _
             As Long

'get key information
Private Declare Function RegQueryInfoKey _
        Lib "advapi32.dll" _
        Alias "RegQueryInfoKeyA" _
            (ByVal hKey As Long, _
             ByVal lpClass As String, _
             lpcbClass As Long, _
             ByVal lpReserved As Long, _
             lpcSubKeys As Long, _
             lpcbMaxSubKeyLen As Long, _
             lpcbMaxClassLen As Long, _
             lpcValues As Long, _
             lpcbMaxValueNameLen As Long, _
             lpcbMaxValueLen As Long, _
             lpcbSecurityDescriptor As Long, _
             lpftLastWriteTime As FILETIME) _
             As Long

'get value information. Note that if you declare the
'lpData parameter as String, you must pass it By Value.
Private Declare Function RegQueryValue _
        Lib "advapi32.dll" _
        Alias "RegQueryValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpValue As String, _
             lpcbValue As Long) _
             As Long
Private Declare Function RegQueryValueEx _
        Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal lpReserved As Long, _
             lpType As Long, _
             lpData As Any, _
             lpcbData As Long) _
             As Long

'replace one key with another
Private Declare Function RegReplaceKey _
        Lib "advapi32.dll" _
        Alias "RegReplaceKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal lpNewFile As String, _
             ByVal lpOldFile As String) _
             As Long

'reads registry information from a file and enters it
'into the registry
Private Declare Function RegRestoreKey _
        Lib "advapi32.dll" _
        Alias "RegRestoreKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             ByVal dwFlags As Long) _
             As Long

'saves a registry key and all its values to a file
Private Declare Function RegSaveKey _
        Lib "advapi32.dll" _
        Alias "RegSaveKeyA" _
            (ByVal hKey As Long, _
             ByVal lpFile As String, _
             lpSecurityAttributes As SECURITY_ATTRIBUTES) _
             As Long

'set the security attributes of the specified registry
'key
Private Declare Function RegSetKeySecurity _
        Lib "advapi32.dll" _
            (ByVal hKey As Long, _
             ByVal SecurityInformation As Long, _
             pSecurityDescriptor As SECURITY_DESCRIPTOR) _
             As Long

'set the information of an existing value. Note that if
'you declare the lpData parameter as String, you must
'pass it By Value.
Private Declare Function RegSetValue _
        Lib "advapi32.dll" _
        Alias "RegSetValueA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String, _
             ByVal dwType As Long, _
             ByVal lpData As String, _
             ByVal cbData As Long) _
             As Long
Private Declare Function RegSetValueEx _
        Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
            (ByVal hKey As Long, _
             ByVal lpValueName As String, _
             ByVal Reserved As Long, _
             ByVal dwType As Long, _
             lpData As Any, _
             ByVal cbData As Long) _
             As Long
             
'unloads a registry key and its values from the registry
Private Declare Function RegUnLoadKey _
        Lib "advapi32.dll" _
        Alias "RegUnLoadKeyA" _
            (ByVal hKey As Long, _
             ByVal lpSubKey As String) _
             As Long

'system information api calls
Private Declare Sub GlobalMemoryStatus _
        Lib "kernel32" _
            (lpBuffer As MEMORYSTATUS)
Private Declare Function GetDiskFreeSpace _
        Lib "kernel32" _
        Alias "GetDiskFreeSpaceA" _
            (ByVal lpRootPathName As String, _
             lpSectorsPerCluster As Long, _
             lpBytesPerSector As Long, _
             lpNumberOfFreeClusters As Long, _
             lpTotalNumberOfClusters As Long) _
             As Long
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () As Long

'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
Public Enum MemType
    CPUUsage
    MemoryUsage
    TotalPhysical
    AvailablePhysical
    TotalPageFile
    AvailablePageFile
    TotalVirtual
    AvailableVirtual
    TotalDisk
    AvailableDisk
End Enum

Public Enum AccessType
    FileInput = 0
    FileOutPut = 1
    FileRandom = 2
    FileBinary = 3
    FileAppend = 4
End Enum

'registry root directory constants
Public Enum RegistryHives
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_CURRENT_USER = &H80000001
    HKEY_DYN_DATA = &H80000006
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_USERS = &H80000003
End Enum

'registry key constants
Public Enum RegistryKeyAccess
    KEY_CREATE_LINK = &H20
    KEY_CREATE_SUB_KEY = &H4
    KEY_ENUMERATE_SUB_KEYS = &H8
    KEY_EVENT = &H1    '  Event contains key event record
    KEY_NOTIFY = &H10
    KEY_QUERY_VALUE = &H1
    KEY_SET_VALUE = &H2
    READ_CONTROL = &H20000
    STANDARD_RIGHTS_ALL = &H1F0000
    STANDARD_RIGHTS_REQUIRED = &HF0000
    Synchronize = &H100000
    STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
    STANDARD_RIGHTS_READ = (READ_CONTROL)
    STANDARD_RIGHTS_WRITE = (READ_CONTROL)
    KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not Synchronize))
    KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not Synchronize))
    KEY_EXECUTE = ((KEY_READ) And (Not Synchronize))
    KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not Synchronize))
End Enum

'registry value attributes
Public Enum RegistryKeyValues
    REG_CREATED_NEW_KEY = &H1               ' New Registry Key created
    REG_EXPAND_SZ = 2            ' Unicode nul terminated string
    REG_FULL_RESOURCE_DESCRIPTOR = 9  ' Resource list in the hardware description
    REG_LINK = 6                ' Symbolic Link (unicode)
    REG_MULTI_SZ = 7             ' Multiple Unicode strings
    REG_NONE = 0                ' No value type
    REG_NOTIFY_CHANGE_ATTRIBUTES = &H2
    REG_NOTIFY_CHANGE_LAST_SET = &H4               ' Time stamp
    REG_NOTIFY_CHANGE_NAME = &H1               ' Create or delete (child)
    REG_NOTIFY_CHANGE_SECURITY = &H8
    REG_OPENED_EXISTING_KEY = &H2               ' Existing Key opened
    REG_OPTION_BACKUP_RESTORE = 4    ' open for backup or restore
    REG_OPTION_CREATE_LINK = 2      ' Created key is a symbolic link
    REG_OPTION_NON_VOLATILE = 0     ' Key is preserved when system is rebooted
    REG_OPTION_RESERVED = 0        ' Parameter is reserved
    REG_OPTION_VOLATILE = 1        ' Key is not preserved when system is rebooted
    REG_REFRESH_HIVE = &H2               ' Unwind changes to last flush
    REG_RESOURCE_LIST = 8          ' Resource list in the resource map
    REG_RESOURCE_REQUIREMENTS_LIST = 10
    REG_SZ = 1                 ' Unicode nul terminated string
    REG_WHOLE_HIVE_VOLATILE = &H1               ' Restore whole hive volatile
    REG_LEGAL_CHANGE_FILTER = (REG_NOTIFY_CHANGE_NAME Or REG_NOTIFY_CHANGE_ATTRIBUTES Or REG_NOTIFY_CHANGE_LAST_SET Or REG_NOTIFY_CHANGE_SECURITY)
    REG_LEGAL_OPTION = (REG_OPTION_RESERVED Or REG_OPTION_NON_VOLATILE Or REG_OPTION_VOLATILE Or REG_OPTION_CREATE_LINK Or REG_OPTION_BACKUP_RESTORE)
End Enum

Public Enum RegistryLongTypes
    REG_BINARY = 3              ' Free form binary
    REG_DWORD = 4               ' 32-bit number
    REG_DWORD_BIG_ENDIAN = 5    ' 32-bit number
    REG_DWORD_LITTLE_ENDIAN = 4 ' 32-bit number (same as REG_DWORD)
End Enum

'error codes returned
Public Enum RegistryErrorCodes
    ERROR_ACCESS_DENIED = 5&
    ERROR_INVALID_PARAMETER = 87 '  dderror
    ERROR_MORE_DATA = 234 '  dderror
    ERROR_SUCCESS = 0&
End Enum

'the shell folders like my documents, recycle bin, temp directory etc.
Public Enum ShellFoldersType
    'registry entry names
    ApplicationDataDir = 0
    TempInetFilesDir = 1
    CookiesDir = 2
    DesktopDir = 3
    FavouritesDir = 4
    FontsDir = 5
    HistoryDir = 6
    LocalAppDataDir = 7
    NetHoodDir = 8
    MyDocumentsDir = 9
    PrintHoodDir = 10
    StartProgramsDir = 11
    RecentDir = 12
    SendToDir = 13
    StartMenuDir = 14
    StartupDir = 15
    TemplatesDir = 16
    
    'these next items are not stored in the registry
    SystemDir = 17
    WindowsDir = 18
    TempDir = 19 'temperory folder is always in the Windows directory
End Enum

Public Enum StartLoginType
    RunBeforeLogin
    RunAfterLogin
End Enum

'the different nt privilages that can be set/unset
Public Enum EnumNTSettings
    'items that can be disabled on the Lock Screen
    CHANGE_PASSWORD = 0
    LOCK_WORKSTATION = 1
    REGISTRY_TOOLS = 2
    TASK_MGR = 3
    
    'the tabs on the Display Properties dialog box
    DISP_APPEARANCE_PAGE = 4
    DISP_BACKGROUND_PAGE = 5
    DISP_CPL = 6
    DISP_SCREENSAVER = 7
    DISP_SETTINGS = 8
End Enum

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
'holds information about the current operating system that the program is
'running on
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

'the current status of physical (ram), virtual memory and the page file.
Public Type MEMORYSTATUS
        dwLength                As Long
        dwMemoryLoad            As Long
        dwTotalPhys             As Long
        dwAvailPhys             As Long
        dwTotalPageFile         As Long
        dwAvailPageFile         As Long
        dwTotalVirtual          As Long
        dwAvailVirtual          As Long
End Type

'defined structures needed
Public Type ACL
        AclRevision             As Byte
        Sbz1                    As Byte
        AclSize                 As Integer
        AceCount                As Integer
        Sbz2                    As Integer
End Type

Public Type FILETIME
        dwLowDateTime           As Long
        dwHighDateTime          As Long
End Type

Public Type SECURITY_DESCRIPTOR
        Revision                As Byte
        Sbz1                    As Byte
        Control                 As Long
        gstrOwner               As Long
        Group                   As Long
        Sacl                    As ACL
        Dacl                    As ACL
End Type

Private Const WIN_INFO_SUBKEY       As String = "Software\Microsoft\Windows\CurrentVersion" 'HKEY_LOCAL_MACHINE
Private Const WIN_NT_INFO_SUBKEY    As String = "Software\Microsoft\Windows NT\CurrentVersion"                              'HKEY_LOCAL_MACHINE
Private Const SHELL_FOLDERS_SUBKEY  As String = ".Default\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders" 'HKEY_USERS
Private Const COUNTRY_SUBKEY        As String = ".Default\Control Panel\International" 'HKEY_USERS
Private Const NT_SETTINGS           As String = WIN_INFO_SUBKEY & "\Policies\System"                                          'HKEY_CURRENT_USER
Private Const W2K_SETTINGS          As String = WIN_INFO_SUBKEY & "\Group Policy Objects\LocalUser\Software\Microsoft\Windows\CurrentVersion\Policies\System"  'HKEY_CURRENT_USER
Private Const STARTUP_AL_SUBKEY     As String = WIN_INFO_SUBKEY & "\Run" 'run after login screen
Private Const STARTUP_BL_SUBKEY     As String = WIN_INFO_SUBKEY & "\RunServices" 'run before login screen

Public Sub CreateSubKey(ByVal enmHive As RegistryHives, _
                        ByVal strSubKey As String)
    
    Dim lngResult As Long
    Dim hKey As Long
    
    'create the key
    lngResult = RegCreateKey(enmHive, _
                             strSubKey & Chr(0), _
                             hKey)
    
    'close the key
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub DeleteSubKey(ByVal enmHive As RegistryHives, _
                        ByVal strSubKey As String)
    'This procedure will delete a key from the registry. Please note that
    'the procedure will not delete key values.
    
    Dim lngResult As Long
    Dim hKey As Long
    
    'open the key
    lngResult = RegOpenKeyEx(enmHive, _
                             strSubKey & Chr(0), _
                             0&, _
                             KEY_ALL_ACCESS, _
                             hKey)
    
    'delete the key
    lngResult = RegDeleteKey(enmHive, hKey)
    
    'close the key
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub DeleteValue(ByVal enmHive As RegistryHives, _
                       ByVal strSubKey As String, _
                       Optional ByVal strEntryLabel As String)
    'This will remove any registry key or entry value
    
    Dim lngResult As Long
    Dim hKey As Long
    Dim strTotalSubKey As String
    
    'create the full registry subkey and entry label
    strTotalSubKey = strSubKey & Chr(0)
    
    'open the subkey/entry
    lngResult = RegOpenKeyEx(enmHive, _
                             strTotalSubKey, _
                             0&, _
                             KEY_ALL_ACCESS, _
                             hKey)
    
    'delete the key/entry from the registry
    lngResult = RegDeleteValue(hKey, strEntryLabel)
    
    'close the handle
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub CreateRegString(ByVal enmHive As RegistryHives, _
                           ByVal strSubKey As String, _
                           ByVal strEntryLabel As String, _
                           ByVal strText As String)
    'This will put some text into the specified key and entry label. This
    'data can be retrieved with the ReadRegString function
    
    Dim lngResult As Long
    Dim hKey As Long
    Dim strTotalSubKey As String
    
    'create a complete sub key and entry path to send to the api call
    strTotalSubKey = strSubKey & Chr(0)
    
    'now create the sub key entry if it does not exist
    lngResult = RegCreateKey(enmHive, strTotalSubKey, hKey)
    
    'if no handle was returned, then exit
    If hKey = 0 Then
        Exit Sub
    End If
    
    'write the text into the key with the specified entry name
    lngResult = RegSetValueEx(hKey, _
                              strEntryLabel, _
                              0&, _
                              REG_SZ, _
                              ByVal strText, _
                              Len(strText))
    
    'close the opened key and exit
    lngResult = RegCloseKey(hKey)
End Sub

Public Function ReadRegString(ByVal enmHive As RegistryHives, _
                              ByVal strSubKey As String, _
                              Optional ByVal strEntry As String) _
                              As String
    'This function will check a registery string entry and
    'return the result.
    
    Dim strText As String
    Dim lngResult As Long
    Dim hOpenKey As Long
    Dim lngBufferSize As Long
    
    'open the registry key
    hOpenKey = GetSubKeyHandle(enmHive, strSubKey)
    
    'check for error
    If hOpenKey = 0 Then
        'return error message
        ReadRegString = "Error : Cannot Open Key"
        Exit Function
    End If
    
    'setup the string to hold the return value
    strText = Space(255)
    lngBufferSize = Len(strText)
    
    'query the information in the key
    lngResult = RegQueryValueEx(hOpenKey, _
                                strEntry, _
                                0, _
                                REG_SZ, _
                                ByVal strText, _
                                lngBufferSize)
    
    'close access to the key
    lngResult = RegCloseKey(hOpenKey)
    
    'check for no values returned
    If Left(strText, 1) = " " Then
        'return error message
        ReadRegString = "Error : Cannot Retrieve String"
        Exit Function
    Else
        'remove the null character
        strText = Left(strText, InStr(1, strText, Chr(0)) - 1)
    End If
    
    'function successful, return owners name
    ReadRegString = strText
End Function

Public Function ReadRegLong(ByVal enmHive As RegistryHives, _
                            ByVal strSubKey As String, _
                            ByVal strEntry As String) _
                            As Long
    'This function will check a registery string
    'entry and return the lngResult.
    
    Dim lngValue As Long
    Dim lngResult As Long
    Dim hOpenKey As Long
    Dim lngBufferSize As Long
    
    'open the registry key
    hOpenKey = GetSubKeyHandle(enmHive, strSubKey)
    
    'check for error
    If hOpenKey = 0 Then
        'return error message
        ReadRegLong = "Error : Cannot Open Key"
        Exit Function
    End If
    
    lngBufferSize = 4
    
    'query the information in the key
    lngResult = RegQueryValueEx(hOpenKey, _
                                strEntry, _
                                ByVal 0&, _
                                REG_BINARY, _
                                lngValue, _
                                lngBufferSize)
    
    'close access to the key
    lngResult = RegCloseKey(hOpenKey)
    
    'function successful, return owners name
    ReadRegLong = lngValue
End Function

Private Function GetSubKeyHandle(ByVal enmHive As RegistryHives, _
                                 ByVal strSubKey As String, _
                                 Optional ByVal enmAccess As RegistryKeyAccess = KEY_READ) _
                                 As Long
    'This function returns a handle to the specified registry key
    
    Dim lngResult As Long
    Dim hKey As Long
    
    'open the registry key
    lngResult = RegOpenKeyEx(enmHive, strSubKey, 0, enmAccess, hKey)
    
    If lngResult <> ERROR_SUCCESS Then
        'could not create key
        hKey = 0
    End If
        
    'return value
    GetSubKeyHandle = hKey
End Function


Public Sub CreateRegLong(ByVal enmHive As RegistryHives, _
                         ByVal strSubKey As String, _
                         ByVal strValueName As String, _
                         ByVal lngData As Long, _
                         Optional ByVal enmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)

    Dim hKey        As Long
    Dim lngResult   As Long
    

    Call CreateSubKey(enmHive, strSubKey)
    
    hKey = GetSubKeyHandle(enmHive, strSubKey, KEY_SET_VALUE)
    
    lngResult = RegSetValueEx(hKey, _
                              strValueName, 0, enmType, lngData, 4)
    
    lngResult = RegCloseKey(hKey)
End Sub

Public Sub OpenVbIdeMaximized(ByVal blnEnable As Boolean)
    
    Const VB_IDE_SUB_KEY    As String = "\Software\Microsoft\Visual Basic\6.0"
    
    Call CreateRegString(HKEY_CURRENT_USER, VB_IDE_SUB_KEY, "MDIMaximized", Trim(str(Abs(blnEnable))))
End Sub




