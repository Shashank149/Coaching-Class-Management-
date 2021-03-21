VERSION 5.00
Begin VB.Form frmAbout1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   8700
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10800
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   PaletteMode     =   2  'Custom
   Picture         =   "frmAbout1.frx":0000
   ScaleHeight     =   6004.895
   ScaleMode       =   0  'User
   ScaleWidth      =   10141.76
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Interval        =   1000
      Left            =   840
      Top             =   3000
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   360
      Top             =   3480
   End
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   240
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   600
      Top             =   1800
   End
   Begin VB.CommandButton cmdBACK 
      BackColor       =   &H0080FFFF&
      Cancel          =   -1  'True
      Caption         =   "<<BACK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   585
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6840
      Width           =   1380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COACHING CLASS "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   810
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   6405
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1095
      Left            =   1920
      TabIndex        =   6
      Top             =   1080
      Width           =   8535
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VERSION 2.0 (BUILD 2008)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   405
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   4020
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "> QUICK RESPONSE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   5400
      TabIndex        =   4
      Top             =   5640
      Width           =   3375
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "> USER FRIENDLY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   450
      Left            =   4440
      TabIndex        =   3
      Top             =   4920
      Width           =   2910
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "> QUICK DATA ENTRY "
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   450
      Left            =   3240
      TabIndex        =   2
      Top             =   4200
      Width           =   3750
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "> EASY INTERFACE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Left            =   2280
      TabIndex        =   1
      Top             =   3480
      Width           =   3150
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   5055
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   8055
   End
End
Attribute VB_Name = "frmAbout1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long



Private Sub cmdok_Click()
  Unload Me
End Sub




Private Sub cmdBACK_Click()
frmAbout1.Hide
frmAbout.Show
End Sub

Private Sub lblVersion_Click()

End Sub

Private Sub Timer1_Timer()
Label3.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer2_Timer()
Label9.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer3_Timer()
Label2.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer4_Timer()
cmdBACK.BackColor = QBColor(Rnd * 15)
End Sub
