VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "FEES COLLECTION.frx":0000
   ScaleHeight     =   7470
   ScaleWidth      =   10710
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3000
      TabIndex        =   19
      Top             =   2760
      Width           =   1695
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   4680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   66715649
      CurrentDate     =   39953
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H80000013&
      Caption         =   "E&xit"
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
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
      TabIndex        =   5
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
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
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
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
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton CMDLAST 
      Caption         =   ">|"
      Height          =   495
      Left            =   6600
      TabIndex        =   12
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   ">>"
      Height          =   495
      Left            =   5160
      TabIndex        =   11
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton CMDPREVIOUS 
      Caption         =   "<<"
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CMDFIRST 
      Caption         =   "|<"
      Height          =   495
      Left            =   2520
      TabIndex        =   9
      Top             =   6480
      Width           =   1335
   End
   Begin VB.TextBox txtstudcap 
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
      Left            =   3000
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
   End
   Begin VB.TextBox txtduration 
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
      Height          =   435
      Left            =   8040
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      Height          =   45
      Left            =   1440
      TabIndex        =   0
      Top             =   5400
      Width           =   7395
   End
   Begin VB.TextBox txtfees 
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
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1320
      Top             =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Course Fees"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1200
      TabIndex        =   18
      Top             =   3360
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Capacity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   4080
      Width           =   1710
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Duration Of Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   5760
      TabIndex        =   16
      Top             =   2640
      Width           =   1980
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Top             =   4680
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000E&
      Height          =   6015
      Left            =   720
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   9375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00400040&
      BackStyle       =   0  'Transparent
      Caption         =   " COURSE ENQUIRY"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   1200
      TabIndex        =   13
      Top             =   2760
      Width           =   1695
   End
End
Attribute VB_Name = "Form8"
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
Combo1.Text = ""
TXTFEES.Text = ""
txtstudcap.Text = ""
DTPicker1.Value = ""
txtduration.Text = ""

Combo1.SetFocus

res.AddNew

 res.Fields(0) = Combo1.Text
 res.Fields(1) = Val(TXTFEES.Text)
 res.Fields(2) = Val(txtstudcap.Text)
 res.Fields(3) = DTPicker1.Value
 res.Fields(4) = Val(txtduration.Text)
 
End Sub

Private Sub cmdsave_Click()
 res.Fields(0) = Combo1.Text
 res.Fields(1) = Val(TXTFEES.Text)
 res.Fields(2) = Val(txtstudcap.Text)
 res.Fields(3) = DTPicker1.Value
 res.Fields(4) = Val(txtduration.Text)
 
 res.Update
 MsgBox "RECORD IS  SUCCESSFULLY ADDED"
 res.MoveLast
End Sub

Private Sub DTPicker1_Change()
DTPicker1.Value = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False")
 res.Open "select * from Enquiry", CN, adOpenDynamic, adLockPessimistic
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


End Sub

Public Sub display()
If res.BOF Then
res.MoveFirst
MsgBox "first record", vbOKOnly, "alert"

End If

Combo1.Text = res.Fields(0)
TXTFEES.Text = res.Fields(1)
txtstudcap.Text = res.Fields(2)
'DTPicker1.Value = res.Fields(3)
txtduration.Text = res.Fields(4)

End Sub

Private Sub List1_Click()
txtcname.Text = List1.Text
End Sub

Private Sub Timer2_Timer()
Label4.ForeColor = QBColor(Rnd * 15)

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

Private Sub txtcname_KeyPress(KeyAscii As Integer)
If IsNumeric(txtcname.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtduration_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtfees_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtstartdate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtstudcap_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub
