VERSION 5.00
Begin VB.Form frmWindows 
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows API's"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FF00FF&
      Caption         =   "Set Back Color"
      Height          =   495
      Left            =   1560
      TabIndex        =   21
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Set active window"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Set Parent"
      Height          =   495
      Left            =   3000
      TabIndex        =   19
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Show Window Normally"
      Height          =   495
      Left            =   1560
      TabIndex        =   18
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00FF0000&
      Caption         =   "Form 2"
      Height          =   1935
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Get Text Length"
      Height          =   495
      Left            =   3000
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Bring To Top"
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Foreground Window"
      Height          =   495
      Left            =   3000
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Get Parent"
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Flash Window"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "y"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Text            =   "height"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "x"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get window caption"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Hide Window"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Maximize Window"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Minimize Window"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Move / Resize Window"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Set Window Caption"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Text            =   "width"
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "frmWindows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim hwnd As Long
Dim length As Long
Dim TheText As String
hwnd = frmWindowsDummy.hwnd
        length = GetWindowTextLength(hwnd) + 1
        TheText = Space(length)
        Call GetWindowText(hwnd, TheText, length)
        Text3.Text = TheText
End Sub

Private Sub Command10_Click()
    BringWindowToTop frmWindowsDummy.hwnd
End Sub

Private Sub Command11_Click()
    a = GetWindowTextLength(frmWindowsDummy.hwnd)
    MsgBox "Text length = " & a, vbInformation, "Text Length"
End Sub

Private Sub Command12_Click()
    ShowWindow frmWindowsDummy.hwnd, SW_SHOWNORMAL
End Sub

Private Sub Command13_Click()
    SetParent frmWindowsDummy.hwnd, frmWindows.hwnd
    MsgBox "Now click 'Show MAXIMIZED"
End Sub

Private Sub Command14_Click()
    SetActiveWindow frmWindowsDummy.hwnd
End Sub

Private Sub Command15_Click()
    SetBkColor frmWindowsDummy.hdc, RGB(255, 0, 255)
End Sub

Private Sub Command2_Click()
    SetWindowText frmWindowsDummy.hwnd, Text2.Text
End Sub

Private Sub Command3_Click()
    FlashWindow frmWindowsDummy.hwnd, True
End Sub

Private Sub Command4_Click()
On Error Resume Next
    MoveWindow frmWindowsDummy.hwnd, Text4.Text, Text1.Text, Text6.Text, Text5.Text, 1
End Sub

Private Sub Command44_Click()
    frmMore.Show
    Me.Hide
    frmWindowsDummy.Hide
End Sub

Private Sub Command5_Click()
    ShowWindow frmWindowsDummy.hwnd, SW_SHOWMINIMIZE
End Sub

Private Sub Command6_Click()
    ShowWindow frmWindowsDummy.hwnd, SW_SHOWMAXIMIZED
End Sub

Private Sub Command7_Click()
    ShowWindow frmWindowsDummy.hwnd, SW_HIDE
End Sub

Private Sub Command8_Click()
If GetParent(frmWindowsDummy.hwnd) = 0 Then
MsgBox "Window has no parent", vbInformation, "Parent"
ElseIf GetParent(frmWindowsDummy.hwnd) = 1 Then
MsgBox "Window has a parent", vbInformation, "Parent"
End If
End Sub

Private Sub Command9_Click()
MsgBox GetForegroundWindow & " as long", vbInformation, "Foreground"
Dim hwnd As Long
Dim length As Long
Dim TheText As String
hwnd = GetForegroundWindow
        length = GetWindowTextLength(hwnd) + 1
        TheText = Space(length)
        Call GetWindowText(hwnd, TheText, length)
MsgBox "Caption " & TheText, vbInformation, "Caption"
End Sub

