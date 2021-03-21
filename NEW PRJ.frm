VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form ADDMISSION 
   Caption         =   "ADMISSION FORM"
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "NEW PRJ.frx":0000
   ScaleHeight     =   9255
   ScaleWidth      =   11355
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox Combo2 
      BackColor       =   &H80000013&
      Height          =   315
      ItemData        =   "NEW PRJ.frx":C4DC
      Left            =   8520
      List            =   "NEW PRJ.frx":C4DE
      TabIndex        =   8
      Top             =   4560
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000013&
      Height          =   315
      Left            =   8520
      TabIndex        =   7
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton CMDFIRST 
      BackColor       =   &H80000013&
      Caption         =   "|<"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton CMDPREVIOUS 
      BackColor       =   &H80000013&
      Caption         =   "<<"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton CMDNEXT 
      Caption         =   ">>"
      Height          =   495
      Left            =   5520
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   15
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton CMDLAST 
      Caption         =   ">|"
      Height          =   495
      Left            =   6960
      TabIndex        =   16
      Top             =   7440
      Width           =   1335
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
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   -2147483629
      CheckBox        =   -1  'True
      Format          =   66781185
      CurrentDate     =   37510
   End
   Begin VB.TextBox txtstudid 
      BackColor       =   &H80000013&
      Height          =   405
      Left            =   3480
      TabIndex        =   0
      Top             =   2160
      Width           =   975
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
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Timer Timer2 
      Interval        =   700
      Left            =   8400
      Top             =   960
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H80000013&
      DataField       =   "SNMAE"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   3135
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   75
      Left            =   1320
      TabIndex        =   17
      Top             =   6240
      Width           =   8535
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtper 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      DataField       =   "PER"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   5
      EndProperty
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
      Height          =   375
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtaddr 
      BackColor       =   &H80000013&
      DataField       =   "ADDRESS"
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
      Height          =   1155
      Left            =   3480
      TabIndex        =   2
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox txtcollege 
      BackColor       =   &H80000013&
      DataField       =   "COLLEGE"
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
      Left            =   3480
      TabIndex        =   5
      Top             =   5400
      Width           =   2295
   End
   Begin VB.TextBox txtphone 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000013&
      DataField       =   "TELPH_NO"
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
      Height          =   375
      Left            =   3480
      MaxLength       =   10
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label8 
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
      Left            =   6720
      TabIndex        =   27
      Top             =   4560
      Width           =   1680
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000018&
      Height          =   6855
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   9855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ADMISSION FORM"
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
      TabIndex        =   26
      Top             =   720
      Width           =   5175
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
      Left            =   1440
      TabIndex        =   25
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Year %"
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
      Left            =   6840
      TabIndex        =   24
      Top             =   2040
      Width           =   1560
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Class"
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
      Left            =   6840
      TabIndex        =   23
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "College"
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
      Left            =   1440
      TabIndex        =   22
      Top             =   5520
      Width           =   930
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
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
      Left            =   1440
      TabIndex        =   21
      Top             =   4920
      Width           =   1275
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
      Left            =   1440
      TabIndex        =   20
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
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
      Left            =   1440
      TabIndex        =   19
      Top             =   4440
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stud Id"
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
      Left            =   1440
      TabIndex        =   18
      Top             =   2280
      Width           =   885
   End
End
Attribute VB_Name = "ADDMISSION"
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
cmddel.SetFocus

End If
End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdnew_Click()
res.MoveLast
txtstudid.Text = ""
txtname.Text = ""
txtaddr.Text = ""
DTPicker1.Value = ""
txtphone.Text = ""
txtcollege.Text = ""
txtper.Text = ""
Combo1.Text = ""
Combo2.Text = ""
txtstudid.SetFocus

res.AddNew

 res.Fields(0) = Val(txtstudid.Text)
 res.Fields(1) = txtname.Text
 res.Fields(2) = txtaddr.Text
 res.Fields(3) = DTPicker1.Value
 res.Fields(4) = Val(txtphone.Text)
 res.Fields(5) = txtcollege.Text
res.Fields(6) = Val(txtper.Text)
res.Fields(7) = Combo1.Text
res.Fields(8) = Combo2.Text

End Sub

Private Sub cmdsave_Click()

 res.Fields(0) = Val(txtstudid.Text)
 res.Fields(1) = txtname.Text
 res.Fields(2) = txtaddr.Text
res.Fields(3) = DTPicker1.Value
 res.Fields(4) = Val(txtphone.Text)
 res.Fields(5) = txtcollege.Text
 res.Fields(6) = Val(txtper.Text)
 res.Fields(7) = Combo1.Text
res.Fields(8) = Combo2.Text

  res.Update
 MsgBox "RECORD IS  SUCCESSFULLY ADDED"
 res.MoveLast
End Sub





Private Sub DTPicker1_Change()
DTPicker1.Value = Format(DTPicker1.Value, "dd/mm/yyyy")
End Sub

Private Sub Form_Load()
CN.Open ("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False")
 res.Open "select * from admisionform", CN, adOpenDynamic, adLockPessimistic
 'rsadd.Open "select * from supplier", cn, adOpenDynamic, adLockPessimistic
 
If res.BOF Then
res.MoveFirst
Else
display
End If

Combo1.AddItem ("Distingtion")
Combo1.AddItem ("First Class")
Combo1.AddItem ("Second Class")
Combo1.AddItem ("Third Class")

Combo2.AddItem ("11th")
Combo2.AddItem ("B-ED")
Combo2.AddItem ("BCA")
Combo2.AddItem ("12th")
Combo2.AddItem ("10th")
Combo2.AddItem ("MCA")
Combo2.AddItem ("MCM")
Combo2.AddItem ("BBA")

End Sub
Public Sub display()
If res.BOF Then
res.MoveFirst
MsgBox "first record", vbOKOnly, "alert"
End If

txtstudid.Text = res.Fields(0)
txtname.Text = res.Fields(1)
txtaddr.Text = res.Fields(2)
DTPicker1.Value = res.Fields(3)
txtphone.Text = res.Fields(4)
txtcollege.Text = res.Fields(5)
txtper.Text = res.Fields(6)
Combo1.Text = res.Fields(7)
Combo2.Text = res.Fields(8)

End Sub

Private Sub List1_Click()
txtcname.Text = List1.Text
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If IsNumeric(Text1.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
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

Private Sub txtaddr_KeyPress(KeyAscii As Integer)
If IsNumeric(txtaddr.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtcname_KeyPress(KeyAscii As Integer)
If IsNumeric(txtcname.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtcollege_KeyPress(KeyAscii As Integer)
If IsNumeric(txtcollege.Text) Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtname_KeyPress(KeyAscii As Integer)
If IsNumeric(txtname.Text) = True Then
MsgBox "ENTER TEXT ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtper_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0

MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtphone_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0


MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub

Private Sub txtstudid_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 58 Or KeyAscii = 8) Then
 KeyAscii = 0

MsgBox "ENTER NUMBER ONLY", vbOKOnly, "ALERT"
End If
End Sub
