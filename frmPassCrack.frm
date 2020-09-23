VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "DUN Password Cracker"
   ClientHeight    =   2760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4605
   ControlBox      =   0   'False
   Icon            =   "frmPassCrack.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2k Emran Hasan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2445
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "ehasan@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      MousePointer    =   10  'Up Arrow
      TabIndex        =   8
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label Error 
      BackStyle       =   0  'Transparent
      Caption         =   "Error Getting Password"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   840
      TabIndex        =   7
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Status : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "R: {Resource including Username} P: {Password}"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "&Clear List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   4
      Top             =   1470
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "&Open DUN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   3
      Top             =   990
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grab Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   2
      Top             =   500
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DUN Password Cracker"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   45
      Width           =   2535
   End
   Begin VB.Image Image5 
      Height          =   300
      Left            =   120
      Picture         =   "frmPassCrack.frx":08CA
      Top             =   480
      Width           =   1425
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   120
      Picture         =   "frmPassCrack.frx":148C
      Top             =   1440
      Width           =   1425
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   120
      Picture         =   "frmPassCrack.frx":204E
      Top             =   960
      Width           =   1425
   End
   Begin VB.Image Close 
      Height          =   240
      Left            =   4320
      Picture         =   "frmPassCrack.frx":2C10
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Min 
      Height          =   240
      Left            =   4080
      Picture         =   "frmPassCrack.frx":2F52
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   2760
      Left            =   0
      Picture         =   "frmPassCrack.frx":3294
      Top             =   0
      Width           =   4590
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
End
End Sub

Private Sub Error_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Form_Load()
Ontop Me
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Label2_Click()
List1.Clear
Call GetPasswords
If List1.Text <> "" Then
Error.Visible = True
Error.ForeColor = vbGreen
Error.Caption = "Got password successfully."
GoTo ends
Else
Error.Visible = True
Error.ForeColor = vbRed
Error.Caption = "Error getting password."
End If
ends:
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbYellow
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack
End Sub

Private Sub Label3_Click()
dun = Shell("C:\WINDOWS\EXPLORER.EXE ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{992CFFA0-F557-101A-88EC-00DD010CCC48}", vbNormalFocus)
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbYellow
Label4.ForeColor = vbBlack
End Sub

Private Sub Label4_Click()
List1.Clear
Error.Visible = False
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbYellow
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Label8_Click()
Shell "c:\program files\internet explorer\iexplore.exe mailto:ehasan@yahoo.com"
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub List1_DblClick()
Dim pik As Integer
pik = List1.ListIndex
def = List1.List(pik)
MsgBox (def)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
Label3.ForeColor = vbBlack
Label4.ForeColor = vbBlack

End Sub

Private Sub Min_Click()
If Form1.Height = 2760 Then
 Form1.Height = 350
Else
 Form1.Height = 2760
End If
End Sub

