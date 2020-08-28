VERSION 5.00
Begin VB.Form Appointment 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Appoinment"
   ClientHeight    =   7080
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15510
   FillColor       =   &H008080FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "APPOINTMENT.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7080
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      DataField       =   "Description"
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
      Height          =   420
      Left            =   9360
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   5040
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      DataField       =   "Appointment date"
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
      Height          =   420
      Left            =   9360
      MousePointer    =   3  'I-Beam
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      DataField       =   "Patient ID"
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
      Height          =   420
      Left            =   9360
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      DataField       =   "Doctor ID"
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
      Height          =   420
      Left            =   9360
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label8 
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
      Height          =   615
      Left            =   1920
      TabIndex        =   11
      Top             =   360
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   360
      Picture         =   "APPOINTMENT.frx":030A
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   13800
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14160
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   2760
      X2              =   13815
      Y1              =   6600
      Y2              =   6615
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   13800
      X2              =   13815
      Y1              =   1680
      Y2              =   6615
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   2760
      X2              =   2775
      Y1              =   1680
      Y2              =   6615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderStyle     =   5  'Dash-Dot-Dot
      X1              =   0
      X2              =   15495
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Label Label6 
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
      Height          =   615
      Left            =   6240
      TabIndex        =   9
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "LAMJUNG HOSPITAL"
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
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Description"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Appointment Date"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   4200
      TabIndex        =   2
      Top             =   3960
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Patient ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   4200
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Doctor's Employee ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      Height          =   4935
      Left            =   2760
      TabIndex        =   10
      Top             =   1680
      Width           =   11055
   End
   Begin VB.Menu Refresh 
      Caption         =   "Refresh"
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
   Begin VB.Menu Back 
      Caption         =   "Back"
   End
End
Attribute VB_Name = "Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "Include your id", vbOKCancel + vbCritical, "Warning"
Else
st = "select doctor_id from appoinment where doctor_id = '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from appoinment where doctor_id = '" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox " Information has been Deleted", vbOKCancel + vbInformation, "Warning"
Else
MsgBox "Doctor ID do not match check your ID again", vbOKCancel + vbCritical, "Warning"
End If
End If
End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Label8_Click()
collection.Show
Unload Me
End Sub

Private Sub Refresh_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Search_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select patient_id,appoinment_date,description from appoinment where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text2.Text = rs.Fields("patient_id")
Text3.Text = rs.Fields("appoinment_date")
Text4.Text = rs.Fields("description")
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
Else
strsq = "select doctor_id from appoinment where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same reg id", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into appoinment (doctor_id,patient_id,appoinment_date,description) values (" _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Text4.Text & "')"
Set rs = cn.Execute(strsql)
Set rs = Nothing
MsgBox "Your information has ben added", vbOKCancel + vbInformation, "warning"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
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
Text1.SetFocus
End If
If Text4.Text = "" Then
Text4.SetFocus
End If
End Sub
