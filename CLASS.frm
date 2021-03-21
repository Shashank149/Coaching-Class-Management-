VERSION 5.00
Begin VB.Form class1 
   Caption         =   "CLASS INFORMATION"
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "CLASS.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton CMDFIRST 
      Caption         =   "|<"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CMDPREVIOUS 
      Caption         =   "<<"
      Height          =   495
      Left            =   2640
      TabIndex        =   10
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   ">>"
      Height          =   495
      Left            =   3960
      TabIndex        =   11
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton CMDLAST 
      BackColor       =   &H80000013&
      Caption         =   ">|"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CommandButton cmdnew 
      BackColor       =   &H80000013&
      Caption         =   "&New"
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
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6000
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
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   770
      Left            =   240
      Top             =   720
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   75
      Left            =   1560
      TabIndex        =   17
      Top             =   5640
      Width           =   4815
   End
   Begin VB.TextBox TXTFEES 
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
      Left            =   3600
      TabIndex        =   4
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox txtcrid 
      BackColor       =   &H80000013&
      DataSource      =   "Adodc1"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtcourse 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.TextBox txtstrn 
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
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label LBLFEES 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FEES"
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
      Left            =   1680
      TabIndex        =   16
      Top             =   4800
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLASS INFORMATION"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   480
      Width           =   5895
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   6015
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   6135
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
      Left            =   1680
      TabIndex        =   14
      Top             =   3600
      Width           =   1680
   End
   Begin VB.Label lbladdr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Strength"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Label lblcid 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class ID"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   3000
      Width           =   1065
   End
End
Attribute VB_Name = "class1"
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


Private Sub cmdnew_Click()
res.MoveLast
txtcrid.Text = ""
txtcourse.Text = ""
txtstrn.Text = ""
txtfees.Text = ""
' txtcid.SetFocus

res.AddNew

 res.Fields(1) = Val(txtcrid.Text)
 res.Fields(2) = txtcourse.Text
 res.Fields(3) = txtstrn.Text
 res.Fields(4) = Val(txtfees.Text)
End Sub

Private Sub cmdsave_Click()
 
 res.Fields(1) = Val(txtcrid.Text)
 res.Fields(2) = txtcourse.Text
 res.Fields(3) = txtstrn.Text
 res.Fields(4) = Val(txtfees.Text)
 res.Update
 MsgBox "RECORD IS  SUCCESSFULLY ADDED"
 res.MoveLast
End Sub


Private Sub Form_Load()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False")

 res.Open "select * from classinfo", CN, adOpenDynamic, adLockPessimistic
 'rsadd.Open "select * from supplier", cn, adOpenDynamic, adLockPessimistic
 

If res.BOF Then
res.MoveFirst
Else
display
End If
End Sub

Private Sub Timer1_Timer()
Label4.ForeColor = QBColor(Rnd * 15)
End Sub
Public Sub display()
If res.BOF Then
res.MoveFirst
MsgBox "first record", vbOKOnly, "alert"
End If

txtcrid.Text = res.Fields(1)
txtcourse.Text = res.Fields(2)
txtstrn.Text = res.Fields(3)
txtfees.Text = res.Fields(4)

End Sub

Private Sub txtcid_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0
MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub




Private Sub txtcourse_KeyPress(KeyAscii As Integer)
If IsNumeric(txtcourse.Text) = True Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtcrid_KeyPress(KeyAscii As Integer)
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

Private Sub txtstrn_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0

MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub
