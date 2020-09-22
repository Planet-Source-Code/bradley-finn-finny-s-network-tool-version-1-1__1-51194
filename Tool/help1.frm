VERSION 5.00
Begin VB.Form help1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   255
   ClientTop       =   390
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Help with Ping"
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
      TabIndex        =   6
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Pinging basicly sends data to a remote computer to determine the speed of the network between the 2 computers."
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label Label3 
      Caption         =   "In the text box labled computer name / IP simply type in the IP or computer name of the computer you want to ping."
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label4 
      Caption         =   "In the text box labled number of pings type the times you want each ping window to ping untill it closes."
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   1440
      Width           =   4455
   End
   Begin VB.Label Label5 
      Caption         =   "In the text box labled number of windows type in how many dos windows you want pinging at once."
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Label Label6 
      Caption         =   "In the text box labled size type in the amont of bytes you want to send on each ping between 1 and 65500."
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   2400
      Width           =   4455
   End
End
Attribute VB_Name = "help1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
