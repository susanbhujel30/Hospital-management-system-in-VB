VERSION 5.00
Begin VB.Form Account 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Section"
   ClientHeight    =   8190
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "ACCOUNT.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   8190
   ScaleWidth      =   15420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
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
      Left            =   13800
      TabIndex        =   35
      Top             =   7200
      Width           =   975
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12840
      MousePointer    =   3  'I-Beam
      TabIndex        =   34
      Text            =   "Text10"
      Top             =   6600
      Width           =   1815
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9600
      MousePointer    =   3  'I-Beam
      TabIndex        =   33
      Text            =   "Text9"
      Top             =   6600
      Width           =   1815
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5280
      MousePointer    =   3  'I-Beam
      TabIndex        =   32
      Text            =   "Text8"
      Top             =   6720
      Width           =   1695
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2160
      MousePointer    =   3  'I-Beam
      TabIndex        =   31
      Text            =   "Text7"
      Top             =   6720
      Width           =   1455
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   12480
      MousePointer    =   3  'I-Beam
      TabIndex        =   30
      Text            =   "Text6"
      Top             =   4680
      Width           =   1815
   End
   Begin VB.OptionButton Option4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "     Bonous        "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11640
      TabIndex        =   28
      Top             =   6480
      Width           =   1335
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Salary "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   27
      Top             =   6480
      Width           =   1215
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
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   12480
      MousePointer    =   3  'I-Beam
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   3600
      Width           =   1815
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
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   12480
      MousePointer    =   3  'I-Beam
      TabIndex        =   19
      Text            =   "Text4"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Bonous   "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   17
      Top             =   6600
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Salary  "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   6600
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   4440
      MousePointer    =   3  'I-Beam
      TabIndex        =   12
      Text            =   "Text3"
      Top             =   4680
      Width           =   1815
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
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   4440
      MousePointer    =   3  'I-Beam
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   4440
      MousePointer    =   3  'I-Beam
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Staff Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   8160
      TabIndex        =   5
      Top             =   2280
      Width           =   6615
      Begin VB.Line Line11 
         X1              =   120
         X2              =   720
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4320
         TabIndex        =   26
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
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
         Left            =   120
         TabIndex        =   25
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   2640
         TabIndex        =   24
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Line Line9 
         X1              =   120
         X2              =   1080
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         BorderStyle     =   2  'Dash
         X1              =   -720
         X2              =   6735
         Y1              =   1800
         Y2              =   1815
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Name"
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
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID"
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
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Doctor Information"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   6615
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   120
         X2              =   615
         Y1              =   4080
         Y2              =   4095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "For"
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
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3840
         TabIndex        =   14
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   13
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2040
         TabIndex        =   11
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Line Line7 
         X1              =   120
         X2              =   1320
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderStyle     =   2  'Dash
         X1              =   0
         X2              =   6600
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor ID"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   1440
         TabIndex        =   6
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Account section"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   14655
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1800
      TabIndex        =   29
      Top             =   360
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   15120
      X2              =   15135
      Y1              =   1560
      Y2              =   7695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   15135
      Y1              =   7680
      Y2              =   7695
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   255
      Y1              =   1560
      Y2              =   7695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   240
      X2              =   15135
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   0
      X2              =   15375
      Y1              =   1320
      Y2              =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Height          =   6135
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   14895
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "ACCOUNT.frx":030A
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
   Begin VB.Menu Submit 
      Caption         =   "Submit"
      Begin VB.Menu Account_doctor 
         Caption         =   "Doctor Account Entry"
      End
      Begin VB.Menu Staff_Account 
         Caption         =   "Staff Account Entry"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "Edit"
      Begin VB.Menu Delete 
         Caption         =   "Delete"
         Begin VB.Menu Doctor_Account_Delete 
            Caption         =   "Doctor Account Delete"
         End
         Begin VB.Menu Staff_Account_Delete 
            Caption         =   "Staff Account Delete"
         End
      End
   End
   Begin VB.Menu Refresh 
      Caption         =   "Refresh"
   End
   Begin VB.Menu search 
      Caption         =   "Search"
      Begin VB.Menu Doctor_Account_Check 
         Caption         =   "Doctor Account Check"
      End
      Begin VB.Menu Staff_Account_Check 
         Caption         =   "Staff Account Check"
      End
   End
End
Attribute VB_Name = "Account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Account_doctor_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
Dim strsq As String

If Text1.Text = "" Then
Text1.SetFocus
ElseIf Text2.Text = "" Then
Text2.SetFocus

ElseIf Text3.Text = "" Then
Text3.SetFocus
Else
strsq = "select doctor_id from doctor_account where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same Reg_id", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into doctor_account(doctor_id,doctor_name,amount,date,salary,bonous) values (" _
& "'" & Text1.Text & "'," _
& "'" & Text2.Text & "'," _
& "'" & Text3.Text & "'," _
& "'" & Label7.Caption & "'," _
& "'" & Text7.Text & "'," _
& "'" & Text8.Text & "')"
Set rs = cn.Execute(strsql)
Set rs = Nothing
MsgBox "Your information has ben added", vbOKCancel + vbInformation, "warning"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text7.Text = ""
Text8.Text = ""
End If
End If
End Sub

Private Sub Close_Click()
Unload Me
collection.Show
End Sub

Private Sub cmdprint_Click()
cmdprint.Visible = False
PrintForm
cmdprint.Visible = True
End Sub

Private Sub Doctor_Account_Check_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select doctor_name,amount,date,salary,bonous from doctor_account where doctor_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text2.Text = rs.Fields("doctor_name")
Text3.Text = rs.Fields("amount")
Label7.Caption = rs.Fields("date")
Text7.Text = rs.Fields("salary")
Text8.Text = rs.Fields("bonous")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

Private Sub Doctor_Account_Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "Include your Id", vbOKCancel + vbCritical, "Warning"
Else
st = "select doctor_id from doctor_account where doctor_id = '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from doctor_account where doctor_id = '" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox " Information has been Deleted", vbOKCancel + vbInformation, "Warning"
Else
MsgBox "Doctor ID do not match check your ID again", vbOKCancel + vbCritical, "Warning"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text7.Text = ""
Text8.Text = ""
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
Label7.Caption = Format(Date, "dd-mm-yyyy")
Label15.Caption = Format(Date, "dd-mm-yyyy")
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
Label7.Caption = Format(Date, "dd-mm-yyyy")
Label15.Caption = Format(Date, "dd-mm-yyyy")
End Sub

Private Sub Staff_Account_Check_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select staff_name,amount,date,salary,bonous from staff_account where staff_id='" & Text4.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Text5.Text = rs.Fields("staff_name")
Text6.Text = rs.Fields("amount")
Label15.Caption = rs.Fields("date")
Text9.Text = rs.Fields("salary")
Text10.Text = rs.Fields("bonous")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

Private Sub Staff_Account_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
Dim strsq As String

If Text4.Text = "" Then
Text4.SetFocus
ElseIf Text5.Text = "" Then
Text5.SetFocus

ElseIf Text6.Text = "" Then
Text6.SetFocus
Else
strsq = "select staff_id from staff_account where staff_id='" & Text4.Text & "'"
Set rs = cn.Execute(strsq)
If Not rs.EOF Then
MsgBox "already have same Reg_id", vbOKCancel + vbCritical, "Warning"
Else
strsql = "insert into staff_account(staff_id,staff_name,amount,date,salary,bonous) values (" _
& "'" & Text4.Text & "'," _
& "'" & Text5.Text & "'," _
& "'" & Text6.Text & "'," _
& "'" & Label15.Caption & "'," _
& "'" & Text9.Text & "'," _
& "'" & Text10.Text & "')"
Set rs = cn.Execute(strsql)
Set rs = Nothing
MsgBox "Your information has ben added", vbOKCancel + vbInformation, "warning"
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text9.Text = ""
Text10.Text = ""
End If
End If
End Sub

Private Sub Staff_Account_Delete_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text4.Text = "" Then
MsgBox "Include your Id", vbOKCancel + vbCritical, "Warning"
Else
st = "select staff_id from staff_account where staff_id = '" & Text4.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from staff_account where staff_id = '" & Text4.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox " Information has been Deleted", vbOKCancel + vbInformation, "Warning"
Else
MsgBox "Staff_ID do not match check your ID again", vbOKCancel + vbCritical, "Warning"
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
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
Text7.SetFocus
End If
If Text3.Text = "" Then
Text3.SetFocus
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
Text4.SetFocus
End If
If Text8.Text = "" Then
Text8.SetFocus
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
Text9.SetFocus
End If
If Text6.Text = "" Then
Text6.SetFocus
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
Text1.SetFocus
End If
If Text10.Text = "" Then
Text10.SetFocus
End If
End Sub
