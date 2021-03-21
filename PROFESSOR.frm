VERSION 5.00
Begin VB.Form Proff 
   Caption         =   "PROFESSOR INFORMATION"
   ClientHeight    =   8535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "PROFESSOR.frx":0000
   ScaleHeight     =   8535
   ScaleWidth      =   9795
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4440
      TabIndex        =   21
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton CMDFIRST 
      Caption         =   "|<"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CMDPREVIOUS 
      Caption         =   "<<"
      Height          =   495
      Left            =   3840
      TabIndex        =   11
      Top             =   7560
      Width           =   1335
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   ">>"
      Height          =   495
      Left            =   5160
      TabIndex        =   12
      Top             =   7560
      Width           =   1455
   End
   Begin VB.CommandButton CMDLAST 
      Caption         =   ">|"
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox txtpid 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmddel 
      BackColor       =   &H80000013&
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   8520
      Top             =   360
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H80000013&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H80000013&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   75
      Left            =   1920
      TabIndex        =   19
      Top             =   6360
      Width           =   6255
   End
   Begin VB.TextBox txtaddr 
      BackColor       =   &H80000013&
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
      Left            =   4440
      TabIndex        =   2
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox txtphno 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtpname 
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4440
      TabIndex        =   1
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox txtsal 
      BackColor       =   &H80000013&
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
      Left            =   4440
      MaxLength       =   7
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label LBLPROF 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROFESSOR INFORMATION INFORMATION"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   20
      Top             =   840
      Width           =   7215
   End
   Begin VB.Label lblprofid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prof ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   18
      Top             =   2640
      Width           =   930
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   17
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   16
      Top             =   4320
      Width           =   990
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   15
      Top             =   5280
      Width           =   795
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   14
      Top             =   3120
      Width           =   1680
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2520
      TabIndex        =   4
      Top             =   3720
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   6375
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   7455
   End
End
Attribute VB_Name = "Proff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CN As New ADODB.Connection
Public res As New ADODB.Recordset


Private Sub cmddel_Click()
res.Delete
If res.EOF = True Then
res.MovePrevious

display
Else
res.MoveNext
display
End If

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub



Private Sub cmdnew_Click()
res.MoveLast
txtpid.Text = ""
Combo1.Text = ""
txtpname.Text = ""
txtaddr.Text = ""
txtphno.Text = ""
txtsal.Text = ""
txtpid.SetFocus

res.AddNew

res.Fields(0) = Val(txtpid.Text)
res.Fields(1) = Combo1.Text
res.Fields(2) = txtpname.Text
res.Fields(3) = txtaddr.Text
res.Fields(4) = txtphno.Text
res.Fields(5) = txtsal.Text

End Sub

Private Sub Command1_Click()

res.Fields(0) = Val(txtpid.Text)
res.Fields(1) = Combo1.Text
res.Fields(2) = txtpname.Text
res.Fields(3) = txtaddr.Text
res.Fields(4) = txtphno.Text
res.Fields(5) = txtsal.Text

res.Update
MsgBox "RECORD IS  SUCCESSFULLY ADDED"
res.MoveLast
End Sub



Private Sub Form_Load()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False")

 res.Open "select * from professorinfo", CN, adOpenDynamic, adLockPessimistic
 'rsadd.Open "select * from supplier", cn, adOpenDynamic, adLockPessimistic
 

If res.BOF Then
res.MoveFirst
Else
display
End If
Combo1.AddItem ("JAVA")
Combo1.AddItem ("C++")
Combo1.AddItem ("C")
Combo1.AddItem ("VB.NET")
Combo1.AddItem ("ASP.NET")
Combo1.AddItem ("AWD")
Combo1.AddItem ("SQL")
Combo1.AddItem ("VB 6.0")
Combo1.AddItem ("Others")
End Sub

Public Sub display()
If res.BOF Then
res.MoveFirst
MsgBox "first record", vbOKOnly, "alert"

End If

txtpid.Text = res.Fields(0)
Combo1.Text = res.Fields(1)
txtpname.Text = res.Fields(2)
txtaddr.Text = res.Fields(3)
txtphno.Text = res.Fields(4)
txtsal.Text = res.Fields(5)


End Sub

Private Sub Timer1_Timer()
LBLPROF.ForeColor = QBColor(Rnd * 15)
End Sub

Private Sub CMDFIRST_Click()
res.MoveFirst
display
End Sub

Private Sub CMDLAST_Click()
res.MoveLast
display
End Sub

Private Sub CMDNEXT_Click()
res.MoveNext
If res.EOF Then
MsgBox "this is last record"
res.MovePrevious
Else

display
End If

End Sub

Private Sub CMDPREVIOUS_Click()
res.MovePrevious
If res.BOF Then
MsgBox "this is first record"
res.MoveNext

Else
display
End If



End Sub

Private Sub txtaddr_KeyPress(KeyAscii As Integer)
If IsNumeric(Tdn.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub



Private Sub txtcid_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
 MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub



Private Sub txtphno_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub



Private Sub txtpid_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtpname_KeyPress(KeyAscii As Integer)
If IsNumeric(txtpname.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub



Private Sub txtsal_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub
