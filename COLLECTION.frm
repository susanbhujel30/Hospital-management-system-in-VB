VERSION 5.00
Begin VB.Form collection 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"COLLECTION.frx":0000
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "COLLECTION.frx":00A8
   MousePointer    =   99  'Custom
   ScaleHeight     =   7395
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   750
      Left            =   13920
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   750
      Left            =   2640
      Top             =   1560
   End
   Begin VB.CommandButton cmd_create 
      Caption         =   "Create New User"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TEST SECTION"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6720
      TabIndex        =   11
      Top             =   6720
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "ACCOUNT SECTION"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   6240
      TabIndex        =   10
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Image Image2 
      Height          =   5055
      Left            =   11640
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   5055
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   3975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Appoinment"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "BILLIG"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DISCHARGE"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   615
      Left            =   6960
      TabIndex        =   6
      Top             =   4920
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "ALL DETAILS SEARCH"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "STAFF ENTRY"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   6480
      TabIndex        =   4
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PATIENT ENTRY"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   3000
      Width           =   3615
   End
   Begin VB.Image Image3 
      Height          =   1320
      Left            =   120
      Picture         =   "COLLECTION.frx":03B2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
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
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   15240
      X2              =   15255
      Y1              =   1560
      Y2              =   7335
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   120
      X2              =   135
      Y1              =   1560
      Y2              =   7455
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   120
      X2              =   15255
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   11520
      X2              =   15135
      Y1              =   7320
      Y2              =   7335
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   15120
      X2              =   15135
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   11520
      X2              =   11535
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00C00000&
      BorderStyle     =   2  'Dash
      X1              =   11520
      X2              =   15135
      Y1              =   2040
      Y2              =   2055
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   480
      X2              =   4815
      Y1              =   7320
      Y2              =   7335
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   480
      X2              =   495
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      X1              =   4680
      X2              =   4695
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   2  'Dash
      Index           =   0
      X1              =   480
      X2              =   4695
      Y1              =   2040
      Y2              =   2055
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5160
      X2              =   11055
      Y1              =   7320
      Y2              =   7335
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5160
      X2              =   5175
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   11040
      X2              =   11055
      Y1              =   2040
      Y2              =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5160
      X2              =   11055
      Y1              =   2040
      Y2              =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Height          =   5295
      Left            =   5160
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Besishahar,Lamjung"
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
      Left            =   6000
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   " LAMJUNG HOSPITAL "
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
      Height          =   735
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Menu Logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "collection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_create_Click()
fill.Show
Unload Me
End Sub



Private Sub Label10_Click()
Account.Show
Unload Me
End Sub

Private Sub Label11_Click()
Unload Me
test.Show
End Sub

Private Sub Label12_Click()

End Sub

Private Sub Label4_Click()
patiententry.Show
Unload Me
End Sub

Private Sub Label5_Click()
staff.Show
Unload Me
End Sub

Private Sub Label6_Click()
detail.Show
Unload Me
End Sub

Private Sub Label7_Click()
Discharge.Show
Unload Me
End Sub

Private Sub Label8_Click()
belling.Show
Unload Me
End Sub

Private Sub Label9_Click()
Appointment.Show
Unload Me
End Sub

Private Sub Logout_Click()
Unload Me
Home.Show
End Sub

Private Sub Timer1_Timer()
Static i
If i = 0 Then
Image1.Picture = LoadPicture("22.jpg")
i = 1
ElseIf i = 1 Then

Image1.Picture = LoadPicture("Untitled-3.jpg")
i = 2
Else
Image1.Picture = LoadPicture("12.jpg")
i = 3
End If
End Sub



Private Sub Timer2_Timer()
Static i
If i = 0 Then
Image2.Picture = LoadPicture("1.jpg")
i = 1
ElseIf i = 1 Then
Image2.Picture = LoadPicture("room1.jpg")
i = 2
Else
Image2.Picture = LoadPicture("23.jpg")
i = 3
End If
End Sub
