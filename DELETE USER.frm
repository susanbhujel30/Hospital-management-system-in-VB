VERSION 5.00
Begin VB.Form delete_user 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete User"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "DELETE USER.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   1665
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
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
      Left            =   2775
      TabIndex        =   2
      Top             =   1065
      Width           =   1020
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Delete"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   870
      TabIndex        =   1
      Top             =   1065
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   3615
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   360
      Width           =   1725
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Password  To Delete "
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
      Height          =   510
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   375
      Width           =   3360
   End
End
Attribute VB_Name = "delete_user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
Home.Show
End Sub

Private Sub cmdOK_Click()
Dim rs As ADODB.Recordset
Dim strsq As String
Dim st As String
If Text1.Text = "" Then
MsgBox "You must include the Password First", vbOKCancel + vbInformation, "Warnng"
Else
st = "select password from login where password= '" & Text1.Text & "'"
Set rs = cn.Execute(st)
If Not rs.EOF Then
strsq = "delete from login where password= '" & Text1.Text & "'"
Set rs = cn.Execute(strsq)
Set rs = Nothing
MsgBox " User has been removed", vbOKCancel + vbInformation, "Warning"
Unload Me
Else
MsgBox "User password do not match", vbOKCancel + vbInformation, "Warning"
End If
End If
End Sub
