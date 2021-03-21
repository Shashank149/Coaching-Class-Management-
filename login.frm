VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   8370
   ClientLeft      =   3045
   ClientTop       =   1305
   ClientWidth     =   9405
   ControlBox      =   0   'False
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      DragIcon        =   "login.frx":B08A
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7200
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   300
      Left            =   0
      Top             =   5880
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   5880
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H80000013&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   2640
      TabIndex        =   0
      Top             =   2520
      Width           =   4095
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter User Name"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   1920
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   3600
      TabIndex        =   10
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " COACHING CLASS MANAGEMENT "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   840
      TabIndex        =   9
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Sylfaen"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   525
      Left            =   2160
      TabIndex        =   7
      Top             =   6240
      Width           =   5040
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   " WELCOME TO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   360
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   9000
      Left            =   -1800
      Picture         =   "login.frx":D82C
      Top             =   -480
      Width           =   12000
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As Integer
Private Sub cmdok_Click()
' for entering correct password
If UCase(Text1.Text) = UCase("shital") Or UCase(Text1.Text) = UCase("punam") Or UCase(Text1.Text) = UCase("pramod") Then
    If UCase(Text2.Text) = UCase("coach") Then
        Timer1.Enabled = True
        ProgressBar1.Visible = True
    Else
        MsgBox "Wrong Password..."
        Text2.Text = ""
        C = C + 1
    End If
Else
    MsgBox "Wrong Login name..."
    Text1.Text = ""
    C = C + 1
End If
If C > 3 Then
    MsgBox "3 Chance Complete...."
    Unload Me
Exit Sub
End If
End Sub
Private Sub Command1_Click()
End
End Sub


Private Sub Timer1_Timer()
'set the timer for progressbar to completing the progerss
If ProgressBar1.Value >= 100 Then
    Unload Me
    MDIForm1.Show
Exit Sub
End If
Label3.Caption = CStr(ProgressBar1.Value) + "% completed"
ProgressBar1 = ProgressBar1 + 1
If ProgressBar1.Value >= 40 And ProgressBar1.Value <= 60 Then
    Label3.Caption = "Loading Supporting Files...."
Else
    If ProgressBar1.Value >= 60 And ProgressBar1.Value <= 90 Then
        Label3.Caption = "Loading Data...."
    Else
    If ProgressBar1.Value >= 100 Then
        Label3.Caption = "Loading Project...."
    End If
    End If
End If
End Sub
Private Sub Timer2_Timer()
Label5.ForeColor = QBColor(15 * Rnd)
End Sub
