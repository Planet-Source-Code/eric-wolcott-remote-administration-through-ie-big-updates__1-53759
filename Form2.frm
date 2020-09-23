VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   5025
      ScaleHeight     =   1170
      ScaleWidth      =   3255
      TabIndex        =   12
      Top             =   1875
      Width           =   3315
      Begin VB.Image Image2 
         Height          =   240
         Left            =   45
         Picture         =   "Form2.frx":0000
         Top             =   60
         Width           =   240
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Assignment Sets USERNAME as 'Admin' and PASSWORD as '1234567'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   30
         TabIndex        =   13
         Top             =   375
         Width           =   3120
      End
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2025
      TabIndex        =   8
      Top             =   2790
      Width           =   2445
   End
   Begin VB.TextBox Text2 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2025
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   2325
      Width           =   2445
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2025
      TabIndex        =   6
      Top             =   1875
      Width           =   2445
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   8535
      TabIndex        =   1
      Top             =   720
      Width           =   8535
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00A76643&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   8535
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PC Online "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   60
         Width           =   6135
      End
   End
   Begin VB.Line Line1 
      X1              =   4740
      X2              =   4740
      Y1              =   1650
      Y2              =   3270
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   420
      TabIndex        =   11
      Top             =   2850
      Width           =   2220
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   420
      TabIndex        =   10
      Top             =   2370
      Width           =   2220
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   420
      TabIndex        =   9
      Top             =   1890
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form2.frx":0264
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6630
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Finished"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7170
      TabIndex        =   4
      Top             =   3450
      Width           =   1290
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6825
      Picture         =   "Form2.frx":031E
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6285
      TabIndex        =   3
      Top             =   840
      Width           =   2190
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Gradiant                        As New clsGradient
Dim cIni                            As New cINIfile
Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
Set Gradiant = New clsGradient
Set cIni = New cINIfile

cIni.Path = App.Path & "\DATA.INI"
cIni.Section = "MAIN"
cIni.Key = "Title"
Me.Caption = cIni.Value
End Sub

Private Sub Form_Resize()
With Gradiant
.Color1 = &HE6A17A
.Color2 = &HD67764
.Angle = 270
.Draw Me
End With

With Gradiant
.Color1 = &HA76643
.Color2 = &HE6A17A
.Angle = 270
.Draw Picture2
End With

End Sub

Private Sub Image1_Click()
If Dofirst <> False Then
Form1.Show
Unload Me
End If
End Sub

Private Sub Label3_Click()
If Dofirst <> False Then
Form1.Show
Unload Me
End If
End Sub

Function Dofirst() As Boolean
If UCase(Text2.Text) <> UCase(Text3.Text) Then MsgBox "Passwords Do not Match": Dofirst = False: Exit Function Else Dofirst = True
If Text1.Text <> "" Then Form1.txtUsername = Text1.Text
If Text2.Text <> "" Then Form1.txtPassword = Text2.Text
End Function
