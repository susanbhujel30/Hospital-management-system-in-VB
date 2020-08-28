VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form doctorentry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Entry"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Upload  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6720
      TabIndex        =   21
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   12480
      TabIndex        =   19
      Text            =   "Text7"
      Top             =   5760
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   12360
      TabIndex        =   15
      Text            =   "Text6"
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   12360
      TabIndex        =   13
      Text            =   "Text5"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   5640
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   2640
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   450
      Left            =   2640
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0FF&
      Height          =   4695
      Left            =   480
      TabIndex        =   3
      Top             =   2160
      Width           =   14535
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   2535
         Left            =   5880
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Label12"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   5280
         TabIndex        =   22
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Education Qualification"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   8280
         TabIndex        =   18
         Top             =   3600
         Width           =   4095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Label9"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   12000
         TabIndex        =   17
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Entry Date"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   9480
         TabIndex        =   16
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   9360
         TabIndex        =   14
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ph No"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   9360
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor_id"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label11 
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
      Left            =   1680
      TabIndex        =   20
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   15240
      X2              =   15255
      Y1              =   1560
      Y2              =   7095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   15255
      Y1              =   7095
      Y2              =   7080
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   15240
      X2              =   240
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   240
      Y1              =   1560
      Y2              =   7080
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0FF&
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   15015
   End
   Begin VB.Line Line1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   0
      X2              =   15480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "DOCTOR ENTRY.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   13920
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Besishahar, Lamjung"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   600
      Width           =   3015
   End
   Begin VB.Label Label18 
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
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Menu Open 
      Caption         =   "Open"
      Begin VB.Menu Close 
         Caption         =   "Close"
         Shortcut        =   ^C
      End
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
End
Attribute VB_Name = "doctorentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
collection.Show
End Sub

Private Sub Command1_Click()
CommonDialog1.FileName = ""
CommonDialog1.Filter = "JPEG|*.jpg|all files|*.*"
CommonDialog1.ShowOpen
Label12 = CommonDialog1.FileName
If Len(Trim(Label15)) < 1 Then
Image2.Picture = LoadPicture(Label12)
End If
End Sub

Private Sub Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "Include your id", vbOKCancel + vbCritical, "Warning"
Else
st = "select doctor_id from doctor_info where doctor_id = '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from doctor_info where doctor_id = '" & Text1.Text & "'"
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
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Label9.Caption = Format(Date, "dd-mm-yyyy")
Label12.Caption = ""
End Sub



Private Sub Search_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select doctor_name,address,gender,ph_no,department,entry_date,education_qualification,image from doctor_info where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text2.Text = rs.Fields("doctor_name")
Text3.Text = rs.Fields("address")
Text4.Text = rs.Fields("gender")
Text5.Text = rs.Fields("ph_no")
Text6.Text = rs.Fields("department")
Label9.Caption = rs.Fields("entry_date")
Text7.Text = rs.Fields("education_qualification")
Label12.Caption = rs.Fields("image")
Image2.Picture = LoadPicture(Label12)
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
Else
strsq = "select doctor_id from doctor_info where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same reg id", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into doctor_info(doctor_id,doctor_name,address,gender,ph_no,department,entry_date,education_qualification,image) values (" _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Text4.Text & "'," _
& "'" & Text5.Text & "'," _
& "'" & Text6.Text & "'," _
& "'" & Label9.Caption & "'," _
& "'" & Text7.Text & "'," _
& "'" & Label12.Caption & "')"
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
Text1.SetFocus
End If
If Text7.Text = "" Then
Text7.SetFocus
End If
End Sub

Private Sub Label11_Click()
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
End Sub
