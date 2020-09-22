VERSION 5.00
Begin VB.Form FrmSome 
   Appearance      =   0  'Flat
   BackColor       =   &H000000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Miscellaneous"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11580
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command47 
      Caption         =   "Set Cursor"
      Height          =   495
      Left            =   3000
      TabIndex        =   118
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton Command46 
      Caption         =   "Email"
      Height          =   495
      Left            =   840
      TabIndex        =   117
      Top             =   7320
      Width           =   1215
   End
   Begin VB.TextBox Text27 
      Height          =   285
      Left            =   240
      TabIndex        =   116
      Text            =   "Address"
      Top             =   6960
      Width           =   2535
   End
   Begin VB.CommandButton Command45 
      Caption         =   "Beep"
      Height          =   495
      Left            =   7320
      TabIndex        =   114
      Top             =   6600
      Width           =   1215
   End
   Begin VB.TextBox Text26 
      Height          =   285
      Left            =   7320
      TabIndex        =   113
      Text            =   "Frequency"
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   7320
      TabIndex        =   112
      Text            =   "Duration"
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command43 
      Caption         =   "Empty Clipboard"
      Height          =   495
      Left            =   240
      TabIndex        =   111
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command41 
      Caption         =   "Caps Lock"
      Height          =   495
      Left            =   10200
      TabIndex        =   109
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command35 
      Caption         =   "Num Lock"
      Height          =   495
      Left            =   10200
      TabIndex        =   108
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Scroll Lock"
      Height          =   495
      Left            =   10200
      TabIndex        =   107
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command44 
      BackColor       =   &H00FF0000&
      Caption         =   "Form 2"
      Height          =   1815
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   5880
      Width           =   1215
   End
   Begin VB.TextBox Text10 
      Height          =   615
      Left            =   8760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   102
      Text            =   "frmmain.frx":0000
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Computer Name"
      Height          =   495
      Left            =   8760
      TabIndex        =   101
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton Command34 
      Caption         =   "Disconnect the Internet"
      Height          =   495
      Left            =   10200
      TabIndex        =   100
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Connect to the Internet"
      Height          =   495
      Left            =   10200
      TabIndex        =   99
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command32 
      Caption         =   "Execute"
      Height          =   495
      Left            =   10200
      TabIndex        =   97
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text17 
      Height          =   765
      Left            =   10200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   96
      Text            =   "frmmain.frx":000E
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command28 
      Caption         =   "Message Box"
      Height          =   495
      Left            =   3000
      TabIndex        =   94
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text14 
      Height          =   285
      Left            =   3000
      TabIndex        =   93
      Text            =   "Caption"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   92
      Text            =   "frmmain.frx":0017
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton Command27 
      Caption         =   "Move File"
      Height          =   495
      Left            =   4440
      TabIndex        =   90
      Top             =   6360
      Width           =   1215
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   4440
      TabIndex        =   89
      Text            =   "From"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   4440
      TabIndex        =   88
      Text            =   "To"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command24 
      Caption         =   "MouseButton L,R"
      Height          =   495
      Left            =   8760
      TabIndex        =   85
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command23 
      Caption         =   "MouseButton R,L"
      Height          =   495
      Left            =   8760
      TabIndex        =   84
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text16 
      Height          =   285
      Left            =   8760
      TabIndex        =   81
      Text            =   "Y"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox Text15 
      Height          =   285
      Left            =   8760
      TabIndex        =   80
      Text            =   "X"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Cursor Position"
      Height          =   495
      Left            =   8760
      TabIndex        =   79
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   8760
      TabIndex        =   78
      Text            =   "Y"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   8760
      TabIndex        =   77
      Text            =   "X"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command25 
      Caption         =   "Cursor Position"
      Height          =   495
      Left            =   8760
      TabIndex        =   76
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text24 
      Height          =   735
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   71
      Text            =   "frmmain.frx":001F
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text23 
      Height          =   735
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   70
      Text            =   "frmmain.frx":0033
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command42 
      Caption         =   "My Computer"
      Height          =   495
      Left            =   7320
      TabIndex        =   69
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text22 
      Height          =   735
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   68
      Text            =   "frmmain.frx":0048
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox Text21 
      Height          =   735
      Left            =   7320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   67
      Text            =   "frmmain.frx":005C
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command40 
      Caption         =   "My Computer"
      Height          =   495
      Left            =   7320
      TabIndex        =   66
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command36 
      Caption         =   "Recycle Bin"
      Height          =   495
      Left            =   5880
      TabIndex        =   61
      Top             =   5040
      Width           =   1215
   End
   Begin VB.TextBox Text20 
      Height          =   735
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   60
      Text            =   "frmmain.frx":0070
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text19 
      Height          =   735
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Text            =   "frmmain.frx":0082
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Recycle Bin"
      Height          =   495
      Left            =   5880
      TabIndex        =   58
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text18 
      Height          =   735
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   57
      Text            =   "frmmain.frx":0094
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   735
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   56
      Text            =   "frmmain.frx":00A7
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command21 
      Caption         =   "Recycle Bin"
      Height          =   495
      Left            =   5880
      TabIndex        =   54
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton Command38 
      Caption         =   "Restart"
      Height          =   495
      Left            =   120
      TabIndex        =   28
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command31 
      Caption         =   "Screensaver OFF"
      Height          =   495
      Left            =   1560
      TabIndex        =   52
      Top             =   6360
      Width           =   1215
   End
   Begin VB.CommandButton Command30 
      Caption         =   "Screensaver ON"
      Height          =   495
      Left            =   1560
      TabIndex        =   53
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   735
      Left            =   4440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   50
      Text            =   "frmmain.frx":00BA
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command22 
      Caption         =   "Shutdown Wallpaper"
      Height          =   495
      Left            =   4440
      TabIndex        =   49
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Notification Area"
      Height          =   495
      Left            =   4440
      TabIndex        =   46
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Notification Area"
      Height          =   495
      Left            =   4440
      TabIndex        =   45
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "CTRL - ALT - DEL"
      Height          =   495
      Left            =   4440
      TabIndex        =   42
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command15 
      Caption         =   "CTRL - ALT - DEL"
      Height          =   495
      Left            =   4440
      TabIndex        =   41
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   855
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Text            =   "frmmain.frx":00C5
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Start Button Text"
      Height          =   495
      Left            =   3000
      TabIndex        =   38
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Desktop"
      Height          =   495
      Left            =   3000
      TabIndex        =   35
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Desktop"
      Height          =   495
      Left            =   3000
      TabIndex        =   34
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Clock"
      Height          =   495
      Left            =   3000
      TabIndex        =   31
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Clock"
      Height          =   495
      Left            =   3000
      TabIndex        =   30
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command39 
      Caption         =   "Logoff"
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command37 
      Caption         =   "Shutdown"
      Height          =   495
      Left            =   120
      TabIndex        =   27
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command20 
      Caption         =   "Cursor"
      Height          =   495
      Left            =   1560
      TabIndex        =   24
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton Command19 
      Caption         =   "Cursor"
      Height          =   495
      Left            =   1560
      TabIndex        =   23
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CD-Drive"
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command17 
      Caption         =   "CD-Drive"
      Height          =   495
      Left            =   120
      TabIndex        =   19
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Task Bar"
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start Button"
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Task Bar"
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start Button"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Button"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      Top             =   360
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   0
      Top             =   6960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   6480
   End
   Begin VB.Label Label45 
      BackStyle       =   0  'Transparent
      Caption         =   "Beep"
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
      Left            =   7320
      TabIndex        =   115
      Top             =   5640
      Width           =   615
   End
   Begin VB.Label Label44 
      BackStyle       =   0  'Transparent
      Caption         =   "On/Off"
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
      Left            =   10200
      TabIndex        =   110
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label43 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3720
      TabIndex        =   105
      Top             =   7920
      Width           =   2775
   End
   Begin VB.Label Label19 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   104
      Top             =   7920
      Width           =   3495
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Computer Name"
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
      Height          =   375
      Left            =   8760
      TabIndex        =   103
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label Label42 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
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
      Left            =   10200
      TabIndex        =   98
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Box"
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
      Left            =   3000
      TabIndex        =   95
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
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
      Left            =   4440
      TabIndex        =   91
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Swap"
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
      Left            =   8760
      TabIndex        =   87
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "Swap"
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
      Left            =   8760
      TabIndex        =   86
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Get"
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
      Left            =   8760
      TabIndex        =   83
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Set"
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
      Left            =   8760
      TabIndex        =   82
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label41 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Name"
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
      Left            =   7320
      TabIndex        =   75
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label40 
      BackStyle       =   0  'Transparent
      Caption         =   "New Name"
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
      Left            =   7320
      TabIndex        =   74
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tip"
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
      Left            =   7320
      TabIndex        =   73
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label38 
      BackStyle       =   0  'Transparent
      Caption         =   "New Tip"
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
      Left            =   7320
      TabIndex        =   72
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label37 
      BackStyle       =   0  'Transparent
      Caption         =   "New Tip"
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
      Left            =   5880
      TabIndex        =   65
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label Label36 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Tip"
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
      Left            =   5880
      TabIndex        =   64
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label35 
      BackStyle       =   0  'Transparent
      Caption         =   "New Name"
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
      Left            =   5880
      TabIndex        =   63
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Name"
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
      Left            =   5880
      TabIndex        =   62
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Empty"
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
      Left            =   5880
      TabIndex        =   55
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "Change"
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
      Left            =   4440
      TabIndex        =   51
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   4440
      TabIndex        =   48
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   4440
      TabIndex        =   47
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   4440
      TabIndex        =   44
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   4440
      TabIndex        =   43
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Rename"
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
      Left            =   3000
      TabIndex        =   40
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   3000
      TabIndex        =   37
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   3000
      TabIndex        =   36
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   3000
      TabIndex        =   33
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   3000
      TabIndex        =   32
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   1560
      TabIndex        =   26
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   1560
      TabIndex        =   25
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
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
      TabIndex        =   22
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
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
      TabIndex        =   21
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      Left            =   1560
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Hide"
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
      TabIndex        =   17
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      Left            =   1560
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
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
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Top"
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
      Left            =   2280
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Left"
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
      Left            =   2280
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label34 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   9480
      TabIndex        =   1
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   6600
      TabIndex        =   0
      Top             =   7920
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private TaskBar As Long, StartButton As Long, Icons As Long, NotificationArea As Long, StartButtonCaption As Long, Clock As Long

Private Sub Command1_Click()
On Error Resume Next
    Call MoveWindow(StartButton, Text3.Text, Text4.Text, Text2.Text, Text1.Text, 1)
End Sub

Private Sub Command10_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
    keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End If
End Sub

Private Sub Command11_Click()
    ShowWindow NotificationArea, 4
End Sub

Private Sub Command12_Click()
    ShowWindow NotificationArea, 0
End Sub

Private Sub Command13_Click()
    ShowWindow Clock, 0
End Sub

Private Sub Command14_Click()
    ShowWindow Clock, 4
End Sub

Private Sub Command15_Click()
Dim ret As Integer
Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, pOld, 0)
End Sub

Private Sub Command16_Click()
Dim ret As Integer
Dim pOld As Boolean
    ret = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, pOld, 0)
End Sub

Private Sub Command17_Click()
    MciSendString "set CDAudio door open", vbNullString, 0, 0
End Sub

Private Sub Command18_Click()
    MciSendString "set CDAudio door closed", vbNullString, 0, 0
End Sub

Private Sub Command19_Click()
    ShowCursor (False)
End Sub

Private Sub Command2_Click()
    CreateRegString HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "", Text18.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"
End Sub

Private Sub Command20_Click()
    ShowCursor (True)
End Sub

Private Sub Command21_Click()
    Call MakeRecycleBinEmpty("C:\", False, False, False)
End Sub

Private Sub Command22_Click()
    Dim desktopwallpaper As String
    desktopwallpaper = Text7.Text
    SystemParametersInfo SPI_SETDESKWALLPAPER, 0, desktopwallpaper, 0
End Sub

Private Sub Command23_Click()
    SwapMouseButton True
End Sub

Private Sub Command24_Click()
    SwapMouseButton False
End Sub

Private Sub Command25_Click()
On Error Resume Next
    SetCursorPos Text8.Text, Text9.Text
End Sub

Private Sub Command26_Click()
    SetComputerName Text10.Text
End Sub

Private Sub Command27_Click()
    MoveFile Text12.Text, Text11.Text
End Sub

Private Sub Command28_Click()
    MessageBox Me.hwnd, Text13.Text, Text14.Text, vbOKOnly
End Sub

Private Sub Command29_Click()
    Dim Result As Long
    Dim Pos As PointAPI
    Result = GetCursorPos(Pos)
    If Result <> 0 Then
    Text15.Text = Pos.X
    Text16.Text = Pos.Y
    Else
    Exit Sub
    End If
End Sub

Private Sub Command3_Click()
    ShowWindow StartButton, 4
End Sub

Private Sub Command30_Click()
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, (True), 0, 0
End Sub

Private Sub Command31_Click()
    SystemParametersInfo SPI_SETSCREENSAVEACTIVE, (False), 0, 0
End Sub

Private Sub Command33_Click()
    InternetAutodial Internet_Autodial_Force_Unattended, 0&
End Sub

Private Sub Command34_Click()
    InternetAutodialHangup (0&)
End Sub

Private Sub Command35_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End If
End Sub

Private Sub Command36_Click()
    CreateRegString HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip", Text20.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"
End Sub

Private Sub Command37_Click()
    ShutDown
End Sub

Private Sub Command38_Click()
    Restart
End Sub

Private Sub Command39_Click()
    LogOff
End Sub

Private Sub Command4_Click()
    ShowWindow TaskBar, 4
End Sub

Private Sub Command40_Click()
    CreateRegString HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip", Text21.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"
End Sub

Private Sub Command41_Click()
Dim Numlockstate As Boolean
Dim caplockstate As Boolean
Dim scrolllockstate As Boolean
Dim keys(0 To 255) As Byte
    Numlockstate = keys(VK_NUMLOCK)
If Numlockstate <> True Then
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End If
End Sub

Private Sub Command42_Click()
    CreateRegString HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "", Text23.Text
    MsgBox "You will have to restart your computer for these changes to take place.", vbInformation + vbOKOnly, "Restart"
End Sub

Private Sub Command43_Click()
    EmptyClipboard
End Sub

Private Sub Command44_Click()
    frmMore.Show
    Me.Hide
End Sub

Private Sub Command45_Click()
On Error Resume Next
    Beep Text26.Text, Text25.Text
End Sub

Private Sub Command46_Click()
    ShellExecute 0&, "Open", "mailto:" & Text27.Text, "", vbNullString, 1
End Sub

Private Sub Command47_Click()
    SetCursor IDC_ARROW
End Sub

Private Sub Command5_Click()
    ShowWindow StartButton, 0
End Sub

Private Sub Command6_Click()
    ShowWindow TaskBar, 0
End Sub

Private Sub Command7_Click()
    ChangeStartButtonText Text5
End Sub

Private Sub Command8_Click()
    ShowWindow Icons, 0
End Sub

Private Sub Command9_Click()
    ShowWindow Icons, 4
End Sub

Private Sub Form_Load()
On Error Resume Next
    TaskBar = FindWindow("Shell_TrayWnd", vbNullString)
    StartButton = FindWindowEx(TaskBar, 0, "button", vbNullString)
    Icons = FindWindowEx(0&, 0&, "Progman", vbNullString)
    NotificationArea = FindWindowEx(TaskBar, 0, "TrayNotifyWnd", vbNullString)
    Clock = FindWindowEx(NotificationArea, 0, "TrayClockWClass", vbNullString)
    StartButtonCaption = GetWindow(TaskBar, 5)
    GetTheCurrentUser
    GetTheComputerName
End Sub

Sub ChangeStartButtonText(txt As TextBox)
On Error Resume Next
    Dim Button As Long
    Dim ShellTrayWnd As Long
    ShellTrayWnd = FindWindow("Shell_TrayWnd", vbNullString)
    Button = FindWindowEx(ShellTrayWnd, 0, "Button", vbNullString)
    Call SendMessageByString(Button, WM_SETTEXT, 0&, txt)
End Sub

Sub GetTheCurrentUser()
On Error Resume Next
    Dim UserNameText As String
    UserNameText = String(200, Chr$(0))
    GetUserName UserNameText, 200
    UserNameText = Left$(UserNameText, InStr(UserNameText, Chr$(0)) - 1)
    Label28.Caption = "The current user is: " & UserNameText
End Sub

Sub GetTheComputerName()
    Dim ComputerNameText As String
    ComputerNameText = String(200, Chr$(0))
    GetComputerName ComputerNameText, 200
    ComputerNameText = Left$(ComputerNameText, InStr(ComputerNameText, Chr$(0)) - 1)
    Label43.Caption = "The computer name is: " & ComputerNameText
End Sub

Private Sub Timer1_Timer()
    Label19.Caption = "Windows has been running for: " & Format(GetTickCount / 60000, "0") & " minutes."
    Label34.Caption = Time
End Sub

Private Sub Timer2_Timer()
    Text6.Text = ReadRegString(HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "")
    Text19.Text = ReadRegString(HKEY_LOCAL_MACHINE, "Software\CLASSES\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}", "InfoTip")
    Text24.Text = ReadRegString(HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "")
    Text22.Text = ReadRegString(HKEY_CLASSES_ROOT, "CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}", "InfoTip")
End Sub
