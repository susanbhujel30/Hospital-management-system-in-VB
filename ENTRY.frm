VERSION 5.00
Begin VB.Form patiententry 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"ENTRY.frx":0000
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15510
   FillColor       =   &H00C00000&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "ENTRY.frx":00A6
   MousePointer    =   99  'Custom
   ScaleHeight     =   7095
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text11 
      DataField       =   "Age"
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
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   21
      Text            =   "Text11"
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox Text10 
      DataField       =   "Status"
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
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   19
      Text            =   "Text10"
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text9 
      DataField       =   "Telephone / mobile number"
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
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   17
      Text            =   "Text9"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      DataField       =   "City"
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
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   15
      Text            =   "Text8"
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox Text7 
      DataField       =   "Registration date"
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
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   1320
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      DataField       =   "Father's / husband's name"
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
      Height          =   375
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   11
      Text            =   "Text6"
      Top             =   5400
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      DataField       =   "Religion"
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
      Height          =   375
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "Text5"
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      DataField       =   "Martial status"
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
      Height          =   375
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Text            =   "Text4"
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      DataField       =   "Patient address"
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
      Height          =   360
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox Text2 
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
      Height          =   375
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      DataField       =   "Reegistration Number"
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
      Height          =   375
      Left            =   3840
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   24
      Text            =   "Text12"
      Top             =   5640
      Width           =   2175
   End
   Begin VB.Label Label16 
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
      Left            =   1920
      TabIndex        =   40
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label30 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   9600
      TabIndex        =   39
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label13 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   38
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label29 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   37
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Label Label28 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11880
      TabIndex        =   36
      Top             =   3480
      Width           =   2775
   End
   Begin VB.Label Label27 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11880
      TabIndex        =   35
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label26 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11640
      TabIndex        =   34
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label Label25 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   33
      Top             =   5760
      Width           =   2655
   End
   Begin VB.Label Label24 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3480
      TabIndex        =   32
      Top             =   5040
      Width           =   2655
   End
   Begin VB.Label Label23 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   31
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label6 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label Label22 
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3360
      TabIndex        =   29
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   28
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "ENTRY.frx":03B0
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
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
   Begin VB.Line Line4 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   135
      Y1              =   1200
      Y2              =   6975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   15375
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      X1              =   15360
      X2              =   15375
      Y1              =   1200
      Y2              =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      X1              =   120
      X2              =   15375
      Y1              =   1200
      Y2              =   1215
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Yrs."
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   14400
      TabIndex        =   22
      Top             =   4800
      Width           =   975
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
      Left            =   6120
      TabIndex        =   26
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
      Left            =   5400
      TabIndex        =   25
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dr.name"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   23
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   15135
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   7680
      TabIndex        =   20
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   7680
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone / Mobile Number"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   7680
      TabIndex        =   16
      Top             =   3000
      Width           =   3255
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   7800
      TabIndex        =   14
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Date"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   7680
      TabIndex        =   12
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Father's / Husband's Name"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   240
      TabIndex        =   10
      Top             =   5160
      Width           =   3735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Religion"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Martial Status"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Paitent Address"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name "
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000003&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient_ID"
      BeginProperty Font 
         Name            =   "Matura MT Script Capitals"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label20 
      Height          =   5775
      Left            =   120
      TabIndex        =   27
      Top             =   1200
      Width           =   15255
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu t 
         Caption         =   "Save as"
         Index           =   1
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu v 
      Caption         =   "View"
      Begin VB.Menu Refresh 
         Caption         =   "Refresh"
      End
      Begin VB.Menu tol 
         Caption         =   "Tools"
         Index           =   2
      End
   End
   Begin VB.Menu e 
      Caption         =   "Edit"
      Begin VB.Menu a 
         Caption         =   "Select all"
      End
      Begin VB.Menu c 
         Caption         =   "Cut"
      End
      Begin VB.Menu co 
         Caption         =   "Copy"
      End
      Begin VB.Menu pa 
         Caption         =   "Paste"
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
      Begin VB.Menu Update 
         Caption         =   "Update"
      End
   End
End
Attribute VB_Name = "patiententry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Close_Click()
Unload Me
End Sub


Private Sub Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "Include your id", vbOKCancel + vbCritical, "Warning"
Else
st = "select patient_id from patient_entry where patient_id= '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from patient_entry where patient_id= '" & Text1.Text & "'"
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
Text11.Text = ""
Text12.Text = ""

End Sub

Private Sub Label16_Click()
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
Text11.Text = ""
Text12.Text = ""
End Sub

Private Sub Search_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select patient_name,patient_address,martial_status,religion,father_or_husband_name,registration_date,city,mb_number,gender,age,dr_name from patient_entry where patient_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text2.Text = rs.Fields("patient_name")
Text3.Text = rs.Fields("patient_address")
Text4.Text = rs.Fields("martial_status")
Text5.Text = rs.Fields("religion")
Text6.Text = rs.Fields("father_or_husband_name")
Text7.Text = rs.Fields("registration_date")
Text8.Text = rs.Fields("city")
Text9.Text = rs.Fields("mb_number")
Text10.Text = rs.Fields("gender")
Text11.Text = rs.Fields("age")
Text12.Text = rs.Fields("dr_name")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

Private Sub Submit_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
Dim strsq As String
With rs
If Text1.Text = "" Then
Label21.Caption = "Registeratiion field is empty"
ElseIf Text2.Text = "" Then
Label22.Caption = "Name field is empty"

ElseIf Text3.Text = "" Then
Label16.Caption = "Address field is empty"

ElseIf Text4.Text = "" Then
Label23.Caption = "status field is empty"

ElseIf Text5.Text = "" Then
Label24.Caption = "Religion field is empty"

ElseIf Text6.Text = "" Then
Label25.Caption = " name field is empty"

ElseIf Text7.Text = "" Then
Label26.Caption = "Date field is empty"

ElseIf Text8.Text = "" Then
Label27.Caption = "city field is empty"

ElseIf Text9.Text = "" Then
Label28.Caption = "Number field is empty"

ElseIf Text10.Text = "" Then
Label29.Caption = "Gender field is empty"

ElseIf Text11.Text = "" Then
Label13.Caption = "age field is empty"
ElseIf Text12.Text = "" Then
Label30.Caption = "DR.name is empty"

Else

strsq = "select patient_id from patient_entry where patient_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same reg entry", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into patient_entry(patient_id,patient_name, patient_address,martial_status,religion,father_or_husband_name,registration_date,city,mb_number,gender,age,dr_name) values (" _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Text4.Text & "'," _
& "'" & Text5.Text & "'," _
& "'" & Text6.Text & "'," _
& "'" & Text7.Text & "'," _
& "'" & Text8.Text & "'," _
& "'" & Text9.Text & "'," _
& "'" & Text10.Text & "'," _
& "'" & Text11.Text & "'," _
& "'" & Text12.Text & "')"
Set rs = cn.Execute(strsql)
Set rs = Nothing
MsgBox "Your information has been added", vbOKCancel + vbInformation, "warning"
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
Text11.Text = ""
Text12.Text = ""

End If
End If
End With

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
Text10.SetFocus
End If
If Text9.Text = "" Then
Text9.SetFocus
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text11.SetFocus
End If
If Text10.Text = "" Then
Text10.SetFocus
End If
End Sub
Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Text12.SetFocus
End If
If Text11.Text = "" Then
Text11.SetFocus
End If
End Sub
Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmd_submit.SetFocus
End If
If Text11.Text = "" Then
Text11.SetFocus
End If
End Sub

Private Sub Update_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
Dim strsq As String
If Text1.Text = "" Then
Label21.Caption = "Registeratiion field is empty"
ElseIf Text2.Text = "" Then
Label22.Caption = "Name field is empty"

ElseIf Text3.Text = "" Then
Label16.Caption = "Address field is empty"

ElseIf Text4.Text = "" Then
Label23.Caption = "status field is empty"

ElseIf Text5.Text = "" Then
Label24.Caption = "Religion field is empty"

ElseIf Text6.Text = "" Then
Label25.Caption = " name field is empty"

ElseIf Text7.Text = "" Then
Label26.Caption = "Date field is empty"

ElseIf Text8.Text = "" Then
Label27.Caption = "city field is empty"

ElseIf Text9.Text = "" Then
Label28.Caption = "Number field is empty"

ElseIf Text10.Text = "" Then
Label29.Caption = "Gender field is empty"

ElseIf Text11.Text = "" Then
Label13.Caption = "age field is empty"
ElseIf Text12.Text = "" Then
Label30.Caption = "DR.name is empty"

Else
strsq = "update patient_entry set patient_name='" & Text2.Text & "'," _
& "patient_address='" & Text3.Text & "'," _
& "martial_status='" & Text4.Text & "'," _
& "religion='" & Text5.Text & "'," _
& "father_or_husband_name='" & Text6.Text & "'," _
& "registration_date='" & Text7.Text & "'," _
& "city='" & Text8.Text & "'," _
& "mb_number='" & Text9.Text & "'," _
& "gender='" & Text10.Text & "'," _
& "age='" & Text11.Text & "'," _
& "dr_name='" & Text12.Text & "'," _
& "where reg_entry='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox "patient information updated", vbOKOnly Or vbExclamation, "warning"
End If
End Sub
