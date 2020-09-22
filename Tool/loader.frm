VERSION 5.00
Begin VB.Form loader 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   120
   ClientTop       =   0
   ClientWidth     =   5100
   FillColor       =   &H80000000&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   0
      Picture         =   "loader.frx":0000
      ScaleHeight     =   1695
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3960
         Top             =   960
      End
   End
End
Attribute VB_Name = "loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer

Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Picture1_Click()
frmMain.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
  counter = counter + 1
If counter = 2 Then
frmMain.Show
Unload Me
End If
End Sub
