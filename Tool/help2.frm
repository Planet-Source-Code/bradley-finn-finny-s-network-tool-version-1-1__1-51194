VERSION 5.00
Begin VB.Form help2 
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Help with Net Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Net send is where you send a message to a remote computer over a network."
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "In the text box labled computer name / IP simply type in the IP or computer name of the computer you want to net send."
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "In the text box labled number of times to send type how many times you want the remote computer to recieve the message."
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   $"help2.frx":0000
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   1920
      Width           =   4455
   End
End
Attribute VB_Name = "help2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
