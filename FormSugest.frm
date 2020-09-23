VERSION 5.00
Begin VB.Form FormSugest 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   1785
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "FormSugest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
List.BackColor = RGB(255, 255, 202)
End Sub

Private Sub List_Click()
Form1.SuggestTag = List.Text
End Sub

Private Sub List_DblClick()
Form1.UseTag = List.Text
End Sub

Private Sub List_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 8, 27, 46 'backcpace, escape, delete
        Form1.UseTag = "" 'use an empty string
    Case 9, 13, 32, 39 'tab,entry, spacebar, right arrow
        Form1.UseTag = List.Text 'accept choice
    Case Else

End Select
End Sub


