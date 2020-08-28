VERSION 5.00
Begin VB.Form contact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"CONTACT.frx":0000
   ClientHeight    =   7095
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   15510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "CONTACT.frx":00A8
   MousePointer    =   99  'Custom
   Picture         =   "CONTACT.frx":03B2
   ScaleHeight     =   7095
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu Back 
      Caption         =   "Back"
   End
End
Attribute VB_Name = "contact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Back_Click()
Home.Show
Unload Me
End Sub
