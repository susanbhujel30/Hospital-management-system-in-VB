VERSION 5.00
Begin VB.Form Gallery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"GALLARY.frx":0000
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MouseIcon       =   "GALLARY.frx":00AA
   MousePointer    =   99  'Custom
   ScaleHeight     =   7095
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
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
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image9 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   11640
      Picture         =   "GALLARY.frx":03B4
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image8 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   8040
      Picture         =   "GALLARY.frx":7C790
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image7 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   4200
      Picture         =   "GALLARY.frx":E9EF5
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image6 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   360
      Picture         =   "GALLARY.frx":1048EA
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Image Image5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   11640
      Picture         =   "GALLARY.frx":106891
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image4 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   8040
      Picture         =   "GALLARY.frx":108C34
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image3 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   4200
      Picture         =   "GALLARY.frx":12AC16
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   360
      Picture         =   "GALLARY.frx":1A4A78
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   120
      Picture         =   "GALLARY.frx":1C7E09
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
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
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   13920
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   1215
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   15240
      X2              =   15255
      Y1              =   1560
      Y2              =   6975
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   15255
      Y1              =   6960
      Y2              =   6975
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   135
      Y1              =   1560
      Y2              =   6975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      X1              =   120
      X2              =   15255
      Y1              =   1560
      Y2              =   1575
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   960
      Width           =   2895
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
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Gallery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()
Home.Show
Unload Me
End Sub
