VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form facebook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facebook"
   ClientHeight    =   7965
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   15480
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6735
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   15375
      ExtentX         =   27120
      ExtentY         =   11880
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   0
      X2              =   15255
      Y1              =   1080
      Y2              =   1095
   End
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "LAMJUNG HOSPITAL"
      BeginProperty Font 
         Name            =   "Rosewood Std Regular"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Besishahar ,lamjung"
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
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   14040
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   14400
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   0
      Picture         =   "FACEBOOK.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1455
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "facebook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_Click()
Unload Me
Home.Show
End Sub

Private Sub Form_Load()
WebBrowser1.Navigate "http://facebook.com"
End Sub

Private Sub Label1_Click()
Unload Me
WebBrowser1.Navigate "http://facebook.com"
Home.Show
End Sub
