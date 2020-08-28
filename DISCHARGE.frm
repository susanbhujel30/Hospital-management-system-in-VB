VERSION 5.00
Begin VB.Form Discharge 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"DISCHARGE.frx":0000
   ClientHeight    =   7395
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "DISCHARGE.frx":0091
   MousePointer    =   99  'Custom
   ScaleHeight     =   7395
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text9 
      DataField       =   "Ward No"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10800
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   4080
      Width           =   3015
   End
   Begin VB.TextBox Text8 
      DataField       =   "Patient Id"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   10800
      MousePointer    =   3  'I-Beam
      TabIndex        =   16
      Text            =   "Text8"
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text7 
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   450
      Left            =   8880
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Text            =   "Text7"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox Text6 
      DataField       =   "Bill no"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      TabIndex        =   14
      Text            =   "Text6"
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      DataField       =   "Time Admitted"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   5160
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      DataField       =   "Appointment"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Text            =   "Text4"
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      DataField       =   "Problem"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Text            =   "Text3"
      Top             =   3600
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "Department"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   2880
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   2040
      TabIndex        =   21
      Top             =   240
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14040
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14400
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Besishahar, Lamjung"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "LaMJUNG HOSPITAL"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5520
      TabIndex        =   18
      Top             =   0
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00004080&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   15015
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00004080&
      BorderStyle     =   2  'Dash
      X1              =   15240
      X2              =   15240
      Y1              =   1320
      Y2              =   7200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   15240
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   240
      Y1              =   1320
      Y2              =   7200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00004080&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   15240
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   8
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Ward No"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Admitted"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   4
      Top             =   5160
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Problem"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Bradley Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Height          =   5895
      Left            =   240
      TabIndex        =   20
      Top             =   1320
      Width           =   15015
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   240
      Picture         =   "DISCHARGE.frx":039B
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu Home 
      Caption         =   "Home"
      Begin VB.Menu Submit 
         Caption         =   "Submit"
      End
      Begin VB.Menu Search 
         Caption         =   "Search"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu Refresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu Back 
      Caption         =   "Back"
   End
End
Attribute VB_Name = "Discharge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text7.Text = "" Then
MsgBox "Include your id", vbOKCancel + vbCritical, "Warning"
Else
st = "select patient_id from discharge where patient_id = '" & Text7.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from discharge where patient_id = '" & Text7.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox " Information has been Deleted", vbOKCancel + vbInformation, "Warning"
Else
MsgBox "patient ID do not match check your ID again", vbOKCancel + vbCritical, "Warning"
End If
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

End Sub

Private Sub Opens_Click()

End Sub

Private Sub Label13_Click()
Unload Me
collection.Show
End Sub

Private Sub Refresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

End Sub
Private Sub Search_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select name,department,problem,appointment,time_admitted,bill_no,ward_no,paid from discharge where patient_id='" & Text7.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text1.Text = rs.Fields("name")
Text2.Text = rs.Fields("department")
Text3.Text = rs.Fields("problem")
Text4.Text = rs.Fields("appointment")
Text5.Text = rs.Fields("time_admitted")
Text6.Text = rs.Fields("bill_no")
Text8.Text = rs.Fields("ward_no")
Text9.Text = rs.Fields("paid")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

Private Sub Submit_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
Dim strsq As String

If Text1.Text = "" Then
Text1.SetFocus
ElseIf Text2.Text = "" Then
Text2.SetFocus

ElseIf Text3.Text = "" Then
Text3.SetFocus

ElseIf Text4.Text = "" Then
Text4.SetFocus

ElseIf Text5.Text = "" Then
Text5.SetFocus

ElseIf Text6.Text = "" Then
Text6.SetFocus

ElseIf Text7.Text = "" Then
Text7.SetFocus

ElseIf Text8.Text = "" Then
Text8.SetFocus

ElseIf Text9.Text = "" Then
Text9.SetFocus
Else
strsq = "select patient_id from discharge where patient_id='" & Text7.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same reg entry", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into discharge(patient_id,name,department,problem,appointment,time_admitted,bill_no,ward_no,paid) values (" _
& "'" & Text7.Text & "'," _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Text4.Text & "'," _
& "'" & Text5.Text & "'," _
& "'" & Text6.Text & "'," _
& "'" & Text8.Text & "'," _
& "'" & Text9.Text & "')"
Set rs = cn.Execute(strsql)
Set rs = Nothing
MsgBox "Your information has ben added", vbOKCancel + vbInformation, "warning"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End If
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text2.SetFocus
End If
If Text1.Text = "" Then
Text1.SetFocus
End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text3.SetFocus
End If
If Text2.Text = "" Then
Text2.SetFocus
End If
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text4.SetFocus
End If
If Text3.Text = "" Then
Text3.SetFocus
End If
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text5.SetFocus
End If
If Text4.Text = "" Then
Text4.SetFocus
End If
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text6.SetFocus
End If
If Text5.Text = "" Then
Text5.SetFocus
End If
End Sub
Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text7.SetFocus
End If
If Text6.Text = "" Then
Text6.SetFocus
End If
End Sub
Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text8.SetFocus
End If
If Text7.Text = "" Then
Text7.SetFocus
End If
End Sub
Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text9.SetFocus
End If
If Text8.Text = "" Then
Text8.SetFocus
End If
End Sub
Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text1.SetFocus
End If
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub


