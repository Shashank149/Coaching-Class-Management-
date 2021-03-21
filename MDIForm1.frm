VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "COACHING CLASS MANAGMENT"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7695
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "4/18/2009"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   20285
            MinWidth        =   20285
            Text            =   "COACHING CLASS MANAGMENT SYSTEM"
            TextSave        =   "COACHING CLASS MANAGMENT SYSTEM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   3528
            MinWidth        =   3528
            TextSave        =   "11:10 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MASTER 
      Caption         =   "MASTER"
      Begin VB.Menu Enquiry 
         Caption         =   "Class Enquiry"
         Shortcut        =   ^C
      End
      Begin VB.Menu STUDENT 
         Caption         =   "STUDENT"
         Shortcut        =   ^S
      End
      Begin VB.Menu PROFFESOR 
         Caption         =   "PROFESSOR"
         Shortcut        =   ^P
      End
      Begin VB.Menu EXIT 
         Caption         =   "EXIT"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu UTILITIES 
      Caption         =   "UTILITIES"
      Begin VB.Menu CALSI 
         Caption         =   "CALCULATOR"
      End
   End
   Begin VB.Menu CLASS 
      Caption         =   "CLASS"
      Begin VB.Menu CLASSINFO 
         Caption         =   "CLASS INFORMATION"
      End
      Begin VB.Menu CLASSFEES1 
         Caption         =   "CLASS FESS"
      End
   End
   Begin VB.Menu ABOUT 
      Caption         =   "ABOUT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABOUT_Click()
frmAbout.Show
End Sub

Private Sub CALSI_Click()
Shell ("calc.exe")
End Sub

Private Sub CLASSFEES1_Click()
FEES.Show
End Sub

Private Sub CLASSINFO_Click()
class1.Show
End Sub

Private Sub Enquiry_Click()
Form8.Show
End Sub

Private Sub EXIT_Click()
Unload Me
End Sub

Private Sub PROFFESOR_Click()
Proff.Show
End Sub

Private Sub STUDENT_Click()
ADDMISSION.Show
End Sub

