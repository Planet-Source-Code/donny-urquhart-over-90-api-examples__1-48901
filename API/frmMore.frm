VERSION 5.00
Begin VB.Form frmMore 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drives"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8220
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Delete Folder"
      Height          =   495
      Left            =   3480
      TabIndex        =   23
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3240
      TabIndex        =   22
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton Command8 
      Caption         =   "File Information"
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3240
      TabIndex        =   20
      Text            =   "FileName"
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Run Dialog Box"
      Height          =   495
      Left            =   3600
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Get The Time"
      Height          =   495
      Left            =   6000
      TabIndex        =   18
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Play"
      Height          =   495
      Left            =   6000
      TabIndex        =   17
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   5160
      TabIndex        =   16
      Text            =   "Filename"
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3240
      TabIndex        =   15
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Create Folder"
      Height          =   495
      Left            =   3480
      TabIndex        =   14
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FF0000&
      Caption         =   "Form3"
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "*.*"
      Top             =   3120
      Width           =   3000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
      Height          =   495
      Left            =   960
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "frmMore.frx":0000
      Left            =   120
      List            =   "frmMore.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3480
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "Directory"
      Top             =   2760
      Width           =   3000
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmMore.frx":0004
      Left            =   120
      List            =   "frmMore.frx":0056
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2160
      Width           =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1800
      Top             =   1920
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "Form1"
      Height          =   1095
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Drive Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   3000
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   3000
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   3000
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3000
   End
End
Attribute VB_Name = "frmMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FrmSome.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim FileCount As Integer
Dim DirCount As Integer
     List1.Clear
     FindFilesInDirectory Text1.Text, Text2.Text, FileCount, DirCount, List1.hwnd
MsgBox "There are: " & FileCount & " files\folders in " & Text1.Text & ".", vbOKOnly, "No of Files"
End Sub

Public Sub GetDiskSpace(ByVal sDrive As String)
On Error Resume Next
    Dim lResult As Long
    Dim lSectorPerCluster As Long
    Dim lBytesPerSector As Long
    Dim lFreeClusters As Long
    Dim lTotalClusters As Long
    
    lResult = GetDiskFreeSpace(sDrive, lSectorPerCluster, lBytesPerSector, lFreeClusters, _
    lTotalClusters)

    m_cDiskSize = CCur(lTotalClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
    m_cDiskFree = CCur(lFreeClusters) * CCur(lSectorPerCluster) * CCur(lBytesPerSector)
    m_cDiskUsed = m_cDiskSize - m_cDiskFree

    If m_cDiskSize <> 0 Then
        m_fFreePercent = m_cDiskFree / m_cDiskSize * 100
    Else
        m_fFreePercent = 0
    End If
    
    Label4.Caption = "There are: " & lSectorPerCluster & " sectors per cluster."
    Label3.Caption = "There are: " & lBytesPerSector & " bytes per sector."
    Label2.Caption = "There are: " & lFreeClusters & " free clusters."
    Label1.Caption = "There are: " & lTotalClusters & " total clusters."
End Sub

Public Sub GetTypeOfDrive(ByVal sDrive As String)
On Error Resume Next
    Select Case GetDriveType(sDrive)
        Case Is = 2
        m_sDriveType = "Removable"
        Case Is = 3
        m_sDriveType = "Fixed"
        Case Is = 4
        m_sDriveType = "Remote"
        Case Is = 5
        m_sDriveType = "CD-Rom"
        Case Is = 6
        m_sDriveType = "RAM Disk"
        Case Else
        m_sDriveType = "Unknown"
    End Select
    Label5.Caption = "The type of drive is: " & m_sDriveType
End Sub

Private Sub Command3_Click()
    Me.Hide
    frmWindows.Show
    frmWindowsDummy.Show
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim security As SECURITY_ATTRIBUTES
    CreateDirectory Text3.Text, security
End Sub

Private Sub Command5_Click()
On Error Resume Next
    Dim filePath As String
    Dim strCmdStr As String
    Dim lngReturnVal As Long
    filePath = Text4.Text
    strCmdStr = "play " & filePath & " fullscreen "
    lngReturnVal = MciSendString(strCmdStr, 0&, 0, 0&)
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim times As SYSTEMTIME
GetSystemTime times
    MsgBox "Date: " & times.wDay, vbInformation, "Time"
    MsgBox "Day: " & times.wDayOfWeek & " of 7", vbInformation, "Time"
    MsgBox "Time Hours: " & times.wHour & " Minutes: " & times.wMinute & " Seconds: " & times.wSecond & " Milliseconds: " & times.wMilliseconds, vbInformation, "Time"
    MsgBox "Year: " & times.wYear, vbInformation, "Time"
    MsgBox "Month: " & times.wMonth, vbInformation, "Time"
End Sub

Private Sub Command7_Click()
On Error Resume Next
    Caption = "Run"
    Description = "Type the name of a program to open, then click OK when finished."
    SHRunDialog Me.hwnd, 0, 0, Caption, Description, 0
End Sub

Private Sub Command8_Click()
On Error Resume Next
Dim info As SHFILEINFO
    SHGetFileInfo Text5.Text, 0&, info, Len(info), SHGFI_DISPLAYNAME
    MsgBox "DisplayName: " & info.szDisplayName
    MsgBox "TypeName: " & info.szTypeName
    MsgBox "Attributes: " & info.dwAttributes
End Sub

Private Sub Command9_Click()
On Error Resume Next
    RemoveDirectory Text6.Text
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    GetDiskSpace Combo1.Text
    GetTypeOfDrive Combo1.Text
End Sub

