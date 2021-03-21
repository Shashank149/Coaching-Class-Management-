VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CCM"
   ClientHeight    =   8400
   ClientLeft      =   2340
   ClientTop       =   2235
   ClientWidth     =   11145
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":B08A
   ScaleHeight     =   5797.83
   ScaleMode       =   0  'User
   ScaleWidth      =   10465.73
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Caption         =   "<<BACK "
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
      Left            =   8880
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7560
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2040
      Top             =   7200
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1200
      Top             =   6960
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   720
      Top             =   5400
   End
   Begin VB.Timer Timer4 
      Interval        =   600
      Left            =   2040
      Top             =   5040
   End
   Begin VB.Label Label1 
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
      Left            =   3480
      TabIndex        =   8
      Top             =   0
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
      Left            =   2520
      TabIndex        =   7
      Top             =   960
      Width           =   8535
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
      Left            =   4560
      TabIndex        =   6
      Top             =   2160
      Width           =   4020
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2008 COE MALEGAON (BK)"
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
      Height          =   735
      Left            =   4320
      TabIndex        =   5
      Top             =   3000
      Width           =   4815
   End
   Begin VB.Label Label4 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Top             =   6600
      Width           =   8175
   End
   Begin VB.Label Label14 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2760
      TabIndex        =   3
      Top             =   5880
      Width           =   6960
   End
   Begin VB.Label Label15 
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
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   6240
      Width           =   8205
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTICE:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   10479.82
      Y1              =   1905.001
      Y2              =   1905.001
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   6960
      Width           =   4575
   End
   Begin VB.Menu G 
      Caption         =   "GENERAL"
      NegotiatePosition=   2  'Middle
   End
   Begin VB.Menu VISH 
      Caption         =   "ABOUT US"
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdBACK_Click()

End Sub

Private Sub CC_Click()

End Sub

Private Sub Command1_Click()
MDIForm1.Show
frmAbout.Hide
End Sub

Private Sub G_Click()
frmAbout.Hide
frmAbout1.Show
End Sub

Private Sub Timer1_Timer()
Label10.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub Timer3_Timer()
Label9.ForeColor = QBColor(Rnd * 15)
'Label9.ForeColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
End Sub

Private Sub Timer4_Timer()
Label16.ForeColor = QBColor(Rnd * 15)
End Sub


Private Sub VISH_Click()
frmAbout.Hide
frmAbout4.Show
End Sub
