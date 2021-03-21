VERSION 5.00
Begin VB.Form frmAbout4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   9735
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   13500
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout4.frx":0000
   ScaleHeight     =   6719.271
   ScaleMode       =   0  'User
   ScaleWidth      =   12677.19
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1200
      Left            =   6720
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   6240
      Top             =   5760
   End
   Begin VB.Timer Timer3 
      Interval        =   1500
      Left            =   6720
      Top             =   1800
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "<<BACK"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      Height          =   735
      Left            =   4800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9000
      Width           =   1575
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">>PRAMOD M."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1005
      Left            =   360
      TabIndex        =   3
      Top             =   6720
      Width           =   5235
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">> PUNAM S."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1005
      Left            =   240
      TabIndex        =   2
      Top             =   4320
      Width           =   4965
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A PROJECT BY>>"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   -1  'True
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1350
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   8340
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   ">> SHITAL J."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   1005
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   5145
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.396
      Y1              =   1697.936
      Y2              =   1697.936
   End
End
Attribute VB_Name = "frmAbout4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
frmAbout4.Hide
frmAbout.Show
End Sub

Private Sub Timer2_Timer()
Command1.BackColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer3_Timer()
Label2.ForeColor = QBColor(Rnd * 15)
End Sub
