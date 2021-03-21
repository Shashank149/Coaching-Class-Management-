VERSION 5.00
Begin VB.Form frmAbout3 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   9390
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   10635
   ClipControls    =   0   'False
   DrawStyle       =   4  'Dash-Dot-Dot
   FillStyle       =   2  'Horizontal Line
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout3.frx":0000
   ScaleHeight     =   6481.145
   ScaleMode       =   0  'User
   ScaleWidth      =   9986.814
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Interval        =   1200
      Left            =   840
      Top             =   5760
   End
   Begin VB.Timer Timer3 
      Interval        =   800
      Left            =   720
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   240
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1320
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0000FFFF&
      Cancel          =   -1  'True
      Caption         =   "OK !!!"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   9240
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8760
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2008 COE MALEGAON (BK)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   3840
      TabIndex        =   9
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " EXTENT POSSIBLE UNDER LAW."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1800
      TabIndex        =   8
      Top             =   8760
      Width           =   3780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CRIMINAL PENELTIES.AND WILL BE PROSECUTED UNDER MAXIMUM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   8400
      Width           =   8175
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
      Left            =   4440
      TabIndex        =   6
      Top             =   600
      Width           =   6405
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MANGMENT SYSTEM"
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1560
      Width           =   7455
   End
   Begin VB.Label Label10 
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
      Left            =   5640
      TabIndex        =   4
      Top             =   2640
      Width           =   4020
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "THESE PROGRAMME IS PROTECTED BY COPYRIGHT LAWS. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   7680
      Width           =   6960
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNAUTHORISED REPRODUCTION MAY RESULT IN SEVERE CIVIL AND "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   8040
      Width           =   8205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   5  'Dash-Dot-Dot
      Index           =   1
      X1              =   0
      X2              =   10029.07
      Y1              =   4058.48
      Y2              =   4058.48
   End
End
Attribute VB_Name = "frmAbout3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
frmAbout3.Hide
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
Label3.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer2_Timer()
Label9.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer3_Timer()
Label11.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer4_Timer()
cmdok.BackColor = QBColor(Rnd * 15)
End Sub
