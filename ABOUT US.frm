VERSION 5.00
Begin VB.Form Aboutus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"ABOUT US.frx":0000
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "ABOUT US.frx":00A9
   MousePointer    =   99  'Custom
   Picture         =   "ABOUT US.frx":03B3
   ScaleHeight     =   7395
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
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
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   120
      Picture         =   "ABOUT US.frx":5A63
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   1800
      Top             =   480
      Width           =   15
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   1215
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   13920
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
   Begin VB.Menu Back 
      Caption         =   "Back"
   End
End
Attribute VB_Name = "Aboutus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label1_Click()
Unload Me
Home.Show
End Sub
