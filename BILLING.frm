VERSION 5.00
Begin VB.Form belling 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"BILLING.frx":0000
   ClientHeight    =   7515
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "BILLING.frx":00AA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7515
   ScaleWidth      =   15540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Print  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14160
      MouseIcon       =   "BILLING.frx":03B4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   11160
      MousePointer    =   3  'I-Beam
      TabIndex        =   21
      Text            =   "Text10"
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   11160
      MousePointer    =   3  'I-Beam
      TabIndex        =   20
      Text            =   "Text9"
      Top             =   4800
      Width           =   2295
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   11160
      MousePointer    =   3  'I-Beam
      TabIndex        =   19
      Text            =   "Text8"
      Top             =   3600
      Width           =   2295
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
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   11160
      MousePointer    =   3  'I-Beam
      TabIndex        =   18
      Text            =   "Text7"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      DataField       =   "OPD Bill no"
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1560
      Width           =   3015
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "Date"
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   3360
      Width           =   2895
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   2
      Text            =   "Text4"
      Top             =   4320
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      DataField       =   "Patient Name"
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Text            =   "Text5"
      Top             =   5280
      Width           =   2895
   End
   Begin VB.TextBox Text6 
      DataField       =   "Paitent details"
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
      ForeColor       =   &H000080FF&
      Height          =   420
      Left            =   3000
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Text            =   "Text6"
      Top             =   6240
      Width           =   2895
   End
   Begin VB.Label Label14 
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
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   2160
      TabIndex        =   23
      Top             =   0
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14520
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14160
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   360
      Picture         =   "BILLING.frx":06BE
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   360
      X2              =   15255
      Y1              =   7200
      Y2              =   7215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   15240
      X2              =   15255
      Y1              =   1320
      Y2              =   7215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   360
      X2              =   375
      Y1              =   1320
      Y2              =   7215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      BorderStyle     =   4  'Dash-Dot
      X1              =   360
      X2              =   15255
      Y1              =   1320
      Y2              =   1335
   End
   Begin VB.Label Label1 
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
      Left            =   5520
      TabIndex        =   17
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "OPD Bill No"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   15
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Consulting Fee"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   12
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   6240
      Width           =   2175
   End
   Begin VB.Label Label8 
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
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   6600
      TabIndex        =   10
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   9
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Concession Amount"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   7
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   6
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000015&
      Height          =   5895
      Left            =   360
      TabIndex        =   22
      Top             =   1320
      Width           =   14895
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
Attribute VB_Name = "belling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdprint_Click()
cmdprint.Visible = False
PrintForm
cmdprint.Visible = True

End Sub

Private Sub Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "Include your id", vbOKCancel + vbCritical, "Warning"
Else
st = "select patient_id from billing where patient_id = '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from billing where patient_id = '" & Text1.Text & "'"
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
Text10.Text = ""
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
End Sub

Private Sub Label14_Click()
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
Text10.Text = ""
End Sub

Private Sub Search_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select opd_bill_no,date,consulting_fee,patient_name,patient_detail,total_amount,concession_amount,net_amount,amount_paid from billing where patient_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text2.Text = rs.Fields("opd_bill_no")
Text3.Text = rs.Fields("date")
Text4.Text = rs.Fields("consulting_fee")
Text5.Text = rs.Fields("patient_name")
Text6.Text = rs.Fields("patient_detail")
Text7.Text = rs.Fields("total_amount")
Text8.Text = rs.Fields("concession_amount")
Text9.Text = rs.Fields("net_amount")
Text10.Text = rs.Fields("amount_paid")
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

ElseIf Text9.Text = "" Then
Text9.SetFocus
Else
strsq = "select patient_id from billing where patient_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same reg id", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into billing (patient_id,opd_bill_no,date,consulting_fee,patient_name,patient_detail,total_amount,concession_amount,net_amount,amount_paid) values (" _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Text4.Text & "'," _
& "'" & Text5.Text & "'," _
& "'" & Text6.Text & "'," _
& "'" & Text7.Text & "'," _
& "'" & Text8.Text & "'," _
& "'" & Text9.Text & "'," _
& "'" & Text10.Text & "')"
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
Text10.Text = ""
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
Text9.Text = Val(Text7.Text) - Val(Text8.Text) - Val(Text4.Text)
Text10.SetFocus
End If
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text10.Text = Val(Text9.Text)

End If
If Text10.Text = "" Then
Text10.SetFocus
End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmd_add.SetFocus
End If
If Text11.Text = "" Then
Text11.SetFocus
End If
End Sub



