VERSION 5.00
Begin VB.Form FEES 
   Caption         =   "Form9"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11580
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   Picture         =   "FEES MASTER.frx":0000
   ScaleHeight     =   8490
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Report"
      Height          =   615
      Left            =   4200
      TabIndex        =   13
      Top             =   7080
      Width           =   2415
   End
   Begin VB.TextBox TXTFEES 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4440
      Width           =   2295
   End
   Begin VB.TextBox txtcourse 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   3240
      Width           =   2295
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
      Height          =   615
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtsid 
      BackColor       =   &H80000013&
      Height          =   375
      Left            =   4920
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin VB.CommandButton CMDLAST 
      Caption         =   ">|"
      Height          =   495
      Left            =   6480
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   ">>"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton CMDPREVIOUS 
      Caption         =   "<<"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton CMDFIRST 
      Caption         =   "|<"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   6360
      Width           =   1335
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
      Height          =   615
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5520
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
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   75
      Left            =   2760
      TabIndex        =   18
      Top             =   5160
      Width           =   4935
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H80000013&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   1680
      Top             =   1080
   End
   Begin VB.Label lblcourse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Id"
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
      Index           =   1
      Left            =   3000
      TabIndex        =   17
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label LBLFEES 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
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
      Index           =   1
      Left            =   3000
      TabIndex        =   16
      Top             =   3240
      Width           =   1725
   End
   Begin VB.Label lblcourse 
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
      Index           =   0
      Left            =   3000
      TabIndex        =   15
      Top             =   3840
      Width           =   1680
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   6135
      Left            =   2040
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Label LBLFEES 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fees"
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
      Index           =   0
      Left            =   3000
      TabIndex        =   14
      Top             =   4440
      Width           =   570
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " FEES MASTER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "FEES"
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
txtsid.Text = ""
txtname.Text = ""
txtcourse.Text = ""
txtfees.Text = ""

res.AddNew

res.Fields(0) = Val(txtsid.Text)
res.Fields(1) = txtname.Text
res.Fields(2) = txtcourse.Text
res.Fields(3) = Val(txtfees.Text)

End Sub


Private Sub cmdsave_Click()

res.Fields(0) = Val(txtsid.Text)
res.Fields(1) = txtname.Text
res.Fields(2) = txtcourse.Text
res.Fields(3) = txtfees.Text

res.Update
 MsgBox "RECORD IS  SUCCESSFULLY ADDED"
 res.MoveLast
End Sub

Private Sub Command1_Click()
DataReport1.Show
End Sub

Private Sub Form_Load()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False")

 res.Open "select * from feesmaster", CN, adOpenDynamic, adLockPessimistic
 'rsadd.Open "select * from supplier", cn, adOpenDynamic, adLockPessimistic
 

If res.BOF Then
res.MoveFirst
Else
display
End If
End Sub

Public Sub display()
If res.BOF Then
res.MoveFirst
MsgBox "first record", vbOKOnly, "alert"

End If

txtsid.Text = res.Fields(0)
txtname.Text = res.Fields(1)
txtcourse.Text = res.Fields(2)
txtfees.Text = res.Fields(3)

End Sub


Private Sub Timer1_Timer()
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




Private Sub txtcourse_KeyPress(KeyAscii As Integer)
If IsNumeric(txtcourse.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub



Private Sub txtfees_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub


Private Sub txtname_KeyPress(KeyAscii As Integer)
If IsNumeric(txtname.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtsid_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub
