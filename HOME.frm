VERSION 5.00
Begin VB.Form home 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"HOME.frx":0000
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15660
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "HOME.frx":00A2
   MousePointer    =   99  'Custom
   ScaleHeight     =   7245
   ScaleWidth      =   15660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   2880
      Top             =   6600
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CLICK HERE TO LOGIN   "
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      UseMaskColor    =   -1  'True
      Width           =   3135
   End
   Begin VB.Image Image4 
      Height          =   855
      Left            =   13920
      Picture         =   "HOME.frx":03AC
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   11880
      Picture         =   "HOME.frx":16EF
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   450
      Left            =   14160
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   1200
   End
   Begin VB.Image Image6 
      Height          =   1815
      Left            =   11760
      Picture         =   "HOME.frx":3CC9
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Info About ME"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3720
      TabIndex        =   11
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Line Line13 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   3360
      X2              =   3375
      Y1              =   1680
      Y2              =   6855
   End
   Begin VB.Line Line12 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   240
      X2              =   3375
      Y1              =   6840
      Y2              =   6855
   End
   Begin VB.Line Line11 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   240
      X2              =   255
      Y1              =   1680
      Y2              =   6855
   End
   Begin VB.Line Line10 
      BorderColor     =   &H8000000D&
      BorderWidth     =   2
      X1              =   240
      X2              =   3375
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderStyle     =   2  'Dash
      X1              =   11640
      X2              =   15135
      Y1              =   4200
      Y2              =   4215
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Visit Again"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   10
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Image Image5 
      Height          =   855
      Left            =   11880
      Picture         =   "HOME.frx":6044B
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   735
      Left            =   13920
      Picture         =   "HOME.frx":7B6B0
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Our Main Gole Is To Give Quality Services To The People. Thank You."
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   6720
      Width           =   8175
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "date"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   12480
      TabIndex        =   8
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "OUR SERVICES"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   480
      TabIndex        =   7
      Top             =   5880
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "NEWS AND EVENTS"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ABUOT US"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3120
      Width           =   1815
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
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1320
      Left            =   120
      Picture         =   "HOME.frx":899A1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1560
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   15255
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   15240
      X2              =   15255
      Y1              =   1680
      Y2              =   6975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   15255
      Y1              =   1680
      Y2              =   1695
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderStyle     =   3  'Dot
      X1              =   11640
      X2              =   11655
      Y1              =   1680
      Y2              =   6975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   15480
      X2              =   15495
      Y1              =   1440
      Y2              =   7095
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   15495
      Y1              =   7080
      Y2              =   7095
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   135
      Y1              =   1440
      Y2              =   7095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   120
      X2              =   15495
      Y1              =   1440
      Y2              =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "GALLERY"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   3135
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
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   7695
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1200
      Left            =   14520
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   450
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmd_add_Click()
Dim rs As ADODB.Recordset
Dim srtsql As String

srtsql = "select username from login where username='" & Text1.Text & "' and password='" & Text2.Text & "' "
Set rs = cn.Execute(srtsql)
With rs
If Not rs.EOF Then
Gallery.Show
MsgBox " information has been granted", vbOKCancel + vbInformation, "Warning"
Else
MsgBox " your informatin is error", vbOKCancel + vbCritical, "Warning"
End If
End With

End Sub



Private Sub Command1_Click()
Login.Show
Unload Me
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label11.Caption = Format(Date, "dd-mm-yyyy")
End Sub




Private Sub Image2_Click()
google.Show
Unload Me
End Sub

Private Sub Image3_Click()
twitter.Show
Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
facebook.Show
End Sub

Private Sub Image5_Click()
Unload Me
youtube.Show
End Sub

Private Sub Label1_Click()
self.Show
End Sub

Private Sub Label10_Click()
service.Show
Unload Me
End Sub



Private Sub Label5_Click()
Gallery.Show
Unload Me
End Sub

Private Sub Label7_Click()
Aboutus.Show
Unload Me
End Sub

Private Sub Label8_Click()
contact.Show
Unload Me
End Sub

Private Sub Label9_Click()
news.Show
Unload Me

End Sub

Private Sub Timer1_Timer()
Label12.Left = Label12.Left + 20
If Label12.Left > Login.ScaleWidth Then
Label12.Left = 0
End If
End Sub
