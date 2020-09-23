VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3030
      Left            =   15
      ScaleHeight     =   2970
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   15
      Width           =   4620
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gradiant                        As New clsGradient
Private Sub Form_Load()
Set cIni = New cINIfile
End Sub

Private Sub Form_Resize()
With Gradiant
.Color1 = &HE6A17A
.Color2 = &HD67764
.Angle = 270
.Draw Picture1
End With
End Sub

