VERSION 5.00
Begin VB.Form detail2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "detail2"
   ClientHeight    =   7695
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MouseIcon       =   "DETAIL2.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   7695
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
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
      Left            =   13920
      TabIndex        =   36
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search"
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
      TabIndex        =   35
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000B&
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
      Left            =   3720
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text2 
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
      Left            =   12120
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   5055
      Left            =   8520
      TabIndex        =   2
      Top             =   1920
      Width           =   6495
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID"
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
         Left            =   1800
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   14
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label21 
         Caption         =   "Last Name"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         TabIndex        =   12
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1440
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1680
         TabIndex        =   10
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label29 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label20 
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
         Left            =   3600
         TabIndex        =   8
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label22 
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
         Height          =   255
         Left            =   3600
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label24 
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
         Left            =   3600
         TabIndex        =   6
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label26 
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
         Left            =   3600
         TabIndex        =   5
         Top             =   3000
         Width           =   2175
      End
      Begin VB.Label Label28 
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
         Left            =   3600
         TabIndex        =   4
         Top             =   3720
         Width           =   2175
      End
      Begin VB.Label Label30 
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
         Left            =   3600
         TabIndex        =   3
         Top             =   4440
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   240
      TabIndex        =   16
      Top             =   1920
      Width           =   7575
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "patient ID"
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
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
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
         Left            =   1080
         TabIndex        =   28
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label7 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   27
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   1320
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label9 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   25
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Reg Date"
         BeginProperty Font 
            Name            =   "Bradley Hand ITC"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1320
         TabIndex        =   24
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label11 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   23
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
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
         Left            =   1320
         TabIndex        =   22
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   21
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
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
         Left            =   1440
         TabIndex        =   20
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label15 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Dr.Name"
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
         Left            =   1200
         TabIndex        =   18
         Top             =   4440
         Width           =   1455
      End
      Begin VB.Label Label17 
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   4440
         Width           =   2415
      End
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
      Left            =   5280
      TabIndex        =   34
      Top             =   120
      Width           =   5535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Besishahar, Lamjung"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   33
      Top             =   720
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   15255
      Y1              =   1320
      Y2              =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   15240
      X2              =   15255
      Y1              =   1320
      Y2              =   7215
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   15255
      Y1              =   7200
      Y2              =   7215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   0
      X2              =   15
      Y1              =   1320
      Y2              =   7215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient info"
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
      Left            =   2520
      TabIndex        =   32
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " Staff Info"
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
      Left            =   10680
      TabIndex        =   31
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   0
      Picture         =   "DETAIL2.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   13800
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14160
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label31 
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
      Left            =   1800
      TabIndex        =   30
      Top             =   240
      Width           =   735
   End
   Begin VB.Menu Refresh 
      Caption         =   "Refresh"
   End
End
Attribute VB_Name = "detail2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select patient_name,patient_address,registration_date,gender,age,dr_name from patient_entry where patient_id='" & Text1.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Label7 = rs.Fields("patient_name")
Label9 = rs.Fields("patient_address")
Label11 = rs.Fields("registration_date")
Label13 = rs.Fields("gender")
Label15 = rs.Fields("age")
Label17 = rs.Fields("dr_name")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

Private Sub Command2_Click()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select first_name,last_name,address,post from staff_info where staff_id='" & Text2.Text & "'"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Label20 = rs.Fields("first_name")
Label22 = rs.Fields("last_name")
Label24 = rs.Fields("address")
Label30 = rs.Fields("post")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub




Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Label31_Click()
Unload Me
Login.Show
End Sub

Private Sub Refresh_Click()
Text1.Text = ""
Text2.Text = ""
Label7.Caption = ""
Label9.Caption = ""
Label11.Caption = ""
Label13.Caption = ""
Label15.Caption = ""
Label17.Caption = ""
Label20.Caption = ""
Label22.Caption = ""
Label24.Caption = ""
Label26.Caption = ""
Label28.Caption = ""
Label30.Caption = ""
End Sub
