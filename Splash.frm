VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   " SPLASH"
   ClientHeight    =   8745
   ClientLeft      =   3375
   ClientTop       =   1740
   ClientWidth     =   9165
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Splash.frx":B08A
   ScaleHeight     =   8745
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   -120
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   500
      Left            =   -120
      Top             =   720
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   -120
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   -120
      Top             =   0
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   240
      X2              =   240
      Y1              =   2520
      Y2              =   7560
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6000
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COACHING CLASS MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000013&
      Height          =   4215
      Left            =   720
      TabIndex        =   1
      Top             =   2880
      Width           =   5655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " VERSION   2.0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   7680
      Width           =   2175
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label2.ForeColor = QBColor(15 * Rnd)
End Sub
Private Sub Timer2_Timer()
' to set the timer
If Timer2.Interval = 5000 Then
    Timer2.Enabled = False
    splash.Hide
Login.Show
End If
End Sub
Private Sub Timer3_Timer()
Line2.BorderColor = QBColor(15 * Rnd)
End Sub
Private Sub Timer4_Timer()
Line1.BorderColor = QBColor(15 * Rnd)
End Sub
