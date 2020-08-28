VERSION 5.00
Begin VB.Form doctor_detail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Doctor Detail"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   17682.11
   ScaleMode       =   0  'User
   ScaleWidth      =   53264.61
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label11 
      Caption         =   "Label11"
      Height          =   3375
      Left            =   12600
      TabIndex        =   10
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "contact_no"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12480
      TabIndex        =   9
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "Label9"
      Height          =   4455
      Left            =   8640
      TabIndex        =   8
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   4935
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   4815
      Left            =   240
      TabIndex        =   6
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Experience"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8760
      TabIndex        =   5
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   15375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   53213.09
      Y1              =   4017.03
      Y2              =   4052.896
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14400
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14040
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "         Lamjung Hospital"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   27.75
         Charset         =   77
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   7695
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1335
      Left            =   0
      Picture         =   "DOCTOR DETAIL.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "          Besishahar, Lamjung"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   840
      Width           =   4095
   End
End
Attribute VB_Name = "doctor_detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Dim strsql As String
strsql = "select doctor_name,doctor_field,experience,contact_no from doctor_list"
Set rs = cn.Execute(strsql)
If Not rs.EOF Then
Label7 = rs.Fields("doctor_name")
Label8 = rs.Fields("doctor_field")
Label9 = rs.Fields("experience")
Label11 = rs.Fields("contact_no")
Else
MsgBox "Record Not Found", vbOKCancel + vbCritical, "Warning"
End If
End Sub

