VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmMain 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Finny's Network tool        Version 1.1"
   ClientHeight    =   7305
   ClientLeft      =   6855
   ClientTop       =   3120
   ClientWidth     =   7425
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   12938
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      BackColor       =   12582912
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ping"
      TabPicture(0)   =   "ping.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "time1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ProgressBar3"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "size1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "windows"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "times(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "host(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "tmrClock1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Net Send"
      TabPicture(1)   =   "ping.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Message"
      Tab(1).Control(1)=   "Label8"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Label5(1)"
      Tab(1).Control(4)=   "Label11(4)"
      Tab(1).Control(5)=   "time2"
      Tab(1).Control(6)=   "Label24"
      Tab(1).Control(7)=   "Label5(2)"
      Tab(1).Control(8)=   "ProgressBar2"
      Tab(1).Control(9)=   "b3"
      Tab(1).Control(10)=   "Command4"
      Tab(1).Control(11)=   "b2"
      Tab(1).Control(12)=   "b1"
      Tab(1).Control(13)=   "Command13"
      Tab(1).Control(14)=   "Command14"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Virtual Drives"
      TabPicture(2)   =   "ping.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label13"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "Label6(2)"
      Tab(2).Control(3)=   "Label3(2)"
      Tab(2).Control(4)=   "Label4(2)"
      Tab(2).Control(5)=   "Label11(3)"
      Tab(2).Control(6)=   "time3"
      Tab(2).Control(7)=   "Label6(4)"
      Tab(2).Control(8)=   "Command8"
      Tab(2).Control(9)=   "v3"
      Tab(2).Control(10)=   "Command7"
      Tab(2).Control(11)=   "v2"
      Tab(2).Control(12)=   "Drive2"
      Tab(2).Control(13)=   "Dir2"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Open shares"
      TabPicture(3)   =   "ping.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label9"
      Tab(3).Control(1)=   "Label1(3)"
      Tab(3).Control(2)=   "Label1(2)"
      Tab(3).Control(3)=   "Label6(1)"
      Tab(3).Control(4)=   "Label11(6)"
      Tab(3).Control(5)=   "time4"
      Tab(3).Control(6)=   "bb"
      Tab(3).Control(7)=   "hs"
      Tab(3).Control(8)=   "host(2)"
      Tab(3).Control(9)=   "Command5"
      Tab(3).Control(10)=   "host(1)"
      Tab(3).Control(11)=   "WebBrowser1"
      Tab(3).ControlCount=   12
      TabCaption(4)   =   "Shares"
      TabPicture(4)   =   "ping.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label11(2)"
      Tab(4).Control(1)=   "time5"
      Tab(4).Control(2)=   "Label3(1)"
      Tab(4).Control(3)=   "Label4(1)"
      Tab(4).Control(4)=   "Label17"
      Tab(4).Control(5)=   "Label15"
      Tab(4).Control(6)=   "Label14"
      Tab(4).Control(7)=   "Label18"
      Tab(4).Control(8)=   "Drive1"
      Tab(4).Control(9)=   "Dir1"
      Tab(4).Control(10)=   "Command10"
      Tab(4).Control(11)=   "c3"
      Tab(4).Control(12)=   "Command9"
      Tab(4).Control(13)=   "c1"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "User Config"
      TabPicture(5)   =   "ping.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Label10"
      Tab(5).Control(1)=   "Label6(0)"
      Tab(5).Control(2)=   "Label1(1)"
      Tab(5).Control(3)=   "Label2(3)"
      Tab(5).Control(4)=   "Label1(5)"
      Tab(5).Control(5)=   "Label2(4)"
      Tab(5).Control(6)=   "Label11(0)"
      Tab(5).Control(7)=   "time6"
      Tab(5).Control(8)=   "Label6(5)"
      Tab(5).Control(9)=   "Label6(6)"
      Tab(5).Control(10)=   "Command6"
      Tab(5).Control(11)=   "deluser"
      Tab(5).Control(12)=   "Command3"
      Tab(5).Control(13)=   "pw"
      Tab(5).Control(14)=   "user"
      Tab(5).Control(15)=   "Command2"
      Tab(5).Control(16)=   "nuser(1)"
      Tab(5).Control(17)=   "npw(1)"
      Tab(5).ControlCount=   18
      TabCaption(6)   =   "FTP"
      TabPicture(6)   =   "ping.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label11(1)"
      Tab(6).Control(1)=   "time7"
      Tab(6).Control(2)=   "lblTransferInfo"
      Tab(6).Control(3)=   "Label1(4)"
      Tab(6).Control(4)=   "Label1(6)"
      Tab(6).Control(5)=   "Label1(7)"
      Tab(6).Control(6)=   "Label1(8)"
      Tab(6).Control(7)=   "Label1(9)"
      Tab(6).Control(8)=   "Label6(3)"
      Tab(6).Control(9)=   "ListView1"
      Tab(6).Control(10)=   "StatusBar1"
      Tab(6).Control(11)=   "ImageList1"
      Tab(6).Control(12)=   "CommonDialog1"
      Tab(6).Control(13)=   "cmdUpload"
      Tab(6).Control(14)=   "cmdDownload"
      Tab(6).Control(15)=   "txtTimeOut"
      Tab(6).Control(16)=   "cmdCancel"
      Tab(6).Control(17)=   "cmdCloseConnection"
      Tab(6).Control(18)=   "cmdQuitSession"
      Tab(6).Control(19)=   "cmdConnect"
      Tab(6).Control(20)=   "txtPassword"
      Tab(6).Control(21)=   "txtUserName"
      Tab(6).Control(22)=   "txtPortNumber"
      Tab(6).Control(23)=   "txtFtpHost"
      Tab(6).Control(24)=   "RichTextBox1"
      Tab(6).ControlCount=   25
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3975
         Left            =   -74400
         TabIndex        =   108
         Top             =   2760
         Width           =   6015
         ExtentX         =   10610
         ExtentY         =   7011
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
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
      Begin VB.TextBox c1 
         Height          =   285
         Left            =   -74280
         TabIndex        =   101
         Top             =   3720
         Width           =   3735
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Share"
         Height          =   495
         Left            =   -69600
         TabIndex        =   100
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox c3 
         Height          =   285
         Left            =   -74280
         TabIndex        =   99
         Top             =   5520
         Width           =   2655
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Unshare"
         Height          =   375
         Left            =   -71520
         TabIndex        =   98
         Top             =   5400
         Width           =   1095
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   -74280
         TabIndex        =   97
         Top             =   2160
         Width           =   4215
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   -74280
         TabIndex        =   96
         Top             =   1440
         Width           =   2055
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   975
         Left            =   -74400
         TabIndex        =   95
         Top             =   5520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1720
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"ping.frx":00C4
      End
      Begin VB.TextBox txtFtpHost 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74400
         TabIndex        =   86
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtPortNumber 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69480
         TabIndex        =   85
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74400
         TabIndex        =   84
         Top             =   1680
         Width           =   2415
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -71880
         PasswordChar    =   "*"
         TabIndex        =   83
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   82
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdQuitSession 
         Caption         =   "&Quit Session"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   81
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdCloseConnection 
         Caption         =   "Clo&se connection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -71280
         TabIndex        =   80
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Ca&ncel operation"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69720
         TabIndex        =   79
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtTimeOut 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -69480
         TabIndex        =   78
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdDownload 
         Caption         =   "&Download..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74400
         TabIndex        =   76
         Top             =   4800
         Width           =   1455
      End
      Begin VB.CommandButton cmdUpload 
         Caption         =   "Uploa&d..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   75
         Top             =   4800
         Width           =   1455
      End
      Begin VB.Timer tmrClock1 
         Interval        =   500
         Left            =   2880
         Top             =   4200
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Enable Net Send"
         Height          =   375
         Left            =   -71280
         TabIndex        =   56
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Disable Net Send"
         Height          =   375
         Left            =   -73560
         TabIndex        =   55
         Top             =   4560
         Width           =   2055
      End
      Begin VB.DirListBox Dir2 
         Height          =   1215
         Left            =   -74160
         TabIndex        =   52
         Top             =   1920
         Width           =   4215
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -74160
         TabIndex        =   51
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox npw 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   44
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox nuser 
         Height          =   285
         Index           =   1
         Left            =   -73920
         TabIndex        =   43
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Create New User"
         Height          =   375
         Left            =   -70560
         TabIndex        =   42
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox user 
         Height          =   285
         Left            =   -73920
         TabIndex        =   41
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox pw 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -72240
         PasswordChar    =   "*"
         TabIndex        =   40
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change Password"
         Height          =   375
         Left            =   -70560
         TabIndex        =   39
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox deluser 
         Height          =   285
         Left            =   -73920
         TabIndex        =   38
         Top             =   5160
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete User"
         Height          =   375
         Left            =   -70560
         TabIndex        =   37
         Top             =   5040
         Width           =   1575
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   1
         Left            =   -74280
         TabIndex        =   32
         Top             =   1440
         Width           =   4335
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Open Share"
         Height          =   495
         Left            =   -69720
         TabIndex        =   31
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   2
         Left            =   -74280
         TabIndex        =   30
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox hs 
         Height          =   285
         Left            =   -72000
         TabIndex        =   29
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CommandButton bb 
         Caption         =   "Open Hidden Share"
         Height          =   495
         Left            =   -69720
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox v2 
         Height          =   285
         Left            =   -74160
         TabIndex        =   24
         Top             =   3480
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Create Virtual Drive"
         Height          =   495
         Left            =   -69600
         TabIndex        =   23
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox v3 
         Height          =   285
         Left            =   -74160
         TabIndex        =   22
         Top             =   5520
         Width           =   3495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Remove Virtual Drive"
         Height          =   495
         Left            =   -70440
         TabIndex        =   21
         Top             =   5280
         Width           =   1335
      End
      Begin VB.TextBox b1 
         Height          =   285
         Left            =   -74040
         TabIndex        =   16
         Top             =   1740
         Width           =   1695
      End
      Begin VB.TextBox b2 
         Height          =   285
         Left            =   -72120
         TabIndex        =   15
         Text            =   "1"
         Top             =   1740
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Send"
         Height          =   375
         Left            =   -69840
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox b3 
         Height          =   285
         Left            =   -74040
         TabIndex        =   12
         Top             =   2340
         Width           =   5295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ping!"
         Height          =   495
         Left            =   5280
         TabIndex        =   6
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox host 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   5
         Top             =   1680
         Width           =   3975
      End
      Begin VB.TextBox times 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   4
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox windows 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Text            =   "1"
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox size1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "65500"
         Top             =   3120
         Width           =   3975
      End
      Begin MSComctlLib.ProgressBar ProgressBar3 
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   3720
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   255
         Left            =   -74040
         TabIndex        =   13
         Top             =   2940
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -70680
         Top             =   3960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -71640
         Top             =   3600
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ping.frx":0146
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ping.frx":049A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ping.frx":07EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "ping.frx":0B42
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar1 
         Height          =   375
         Left            =   -74400
         TabIndex        =   77
         Top             =   6600
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   1
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   2175
         Left            =   -74400
         TabIndex        =   87
         Top             =   2520
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3836
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Delete User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   -73920
         TabIndex        =   112
         Top             =   4200
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Change User Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   -73920
         TabIndex        =   111
         Top             =   2520
         Width           =   4935
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Messenger Service Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74040
         TabIndex        =   110
         Top             =   3960
         Width           =   5295
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Remove Virtual Drives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   -74160
         TabIndex        =   109
         Top             =   4440
         Width           =   5055
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Unshare Shares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   107
         Top             =   4560
         Width           =   6615
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "Create Shares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   106
         Top             =   720
         Width           =   6615
      End
      Begin VB.Label Label15 
         Caption         =   "Name of new share (eg Games)"
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
         Left            =   -74280
         TabIndex        =   105
         Top             =   3480
         Width           =   2775
      End
      Begin VB.Label Label17 
         Caption         =   "Name of share to be unshared"
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
         Left            =   -74280
         TabIndex        =   104
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   103
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   102
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "FTP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   -73920
         TabIndex        =   94
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&FTP remote host:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   -74400
         TabIndex        =   93
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Port number:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   -69480
         TabIndex        =   92
         Top             =   840
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&User name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   -74400
         TabIndex        =   91
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pass&word:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   -71880
         TabIndex        =   90
         Top             =   1440
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Time out:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   -69480
         TabIndex        =   89
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label lblTransferInfo 
         Alignment       =   2  'Center
         Caption         =   "lblTransferInfo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74385
         TabIndex        =   88
         Top             =   5280
         Width           =   6120
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Click me for Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72600
         TabIndex        =   74
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Click me for Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2400
         TabIndex        =   73
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label time7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74760
         TabIndex        =   72
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   71
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label time1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   7080
         Width           =   1575
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   6
         Left            =   -70440
         TabIndex        =   65
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   5
         Left            =   4560
         TabIndex        =   64
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   4
         Left            =   -70440
         TabIndex        =   63
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   3
         Left            =   -70440
         TabIndex        =   62
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   2
         Left            =   -70440
         TabIndex        =   61
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   1
         Left            =   -70440
         TabIndex        =   60
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Made by Finny   finnyrulz@hotmail.com"
         Height          =   255
         Index           =   0
         Left            =   -70440
         TabIndex        =   59
         Top             =   7080
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "New Password"
         Height          =   255
         Index           =   4
         Left            =   -72240
         TabIndex        =   58
         Top             =   3120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "User Name"
         Height          =   255
         Index           =   5
         Left            =   -73920
         TabIndex        =   57
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Select Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   54
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select drive:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   -74160
         TabIndex        =   53
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Password"
         Height          =   255
         Index           =   3
         Left            =   -71880
         TabIndex        =   50
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "User Name"
         Height          =   255
         Index           =   1
         Left            =   -73560
         TabIndex        =   49
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Existing username"
         Height          =   255
         Index           =   0
         Left            =   -74040
         TabIndex        =   48
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "New password"
         Height          =   255
         Index           =   0
         Left            =   -72240
         TabIndex        =   47
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Create New User"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   -73920
         TabIndex        =   46
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Username to delete"
         Height          =   255
         Left            =   -73920
         TabIndex        =   45
         Top             =   4920
         Width           =   3135
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Open Shares"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74400
         TabIndex        =   36
         Top             =   600
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Computername or IP"
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
         Index           =   2
         Left            =   -74280
         TabIndex        =   35
         Top             =   1200
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Computername or IP"
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
         Index           =   3
         Left            =   -74280
         TabIndex        =   34
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Hidden Share Name"
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
         Left            =   -72000
         TabIndex        =   33
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Virtual Drives"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -74160
         TabIndex        =   27
         Top             =   660
         Width           =   5055
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Caption         =   "Vitual Drive Letter (eg F)"
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
         Left            =   -74160
         TabIndex        =   26
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label13 
         Caption         =   "Virtual Drive letter to be removed (eg F)"
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
         Left            =   -74160
         TabIndex        =   25
         Top             =   5280
         Width           =   3855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Net Send"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -74280
         TabIndex        =   20
         Top             =   660
         Width           =   5295
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Computer Name"
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
         Left            =   -74040
         TabIndex        =   19
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Number of times to send"
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
         Left            =   -72240
         TabIndex        =   18
         Top             =   1500
         Width           =   2295
      End
      Begin VB.Label Message 
         Alignment       =   2  'Center
         Caption         =   "Message"
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
         Left            =   -74040
         TabIndex        =   17
         Top             =   2100
         Width           =   5055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Number of Pings"
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
         Index           =   0
         Left            =   3120
         TabIndex        =   11
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Computername or IP"
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
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   1440
         Width           =   3975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Number of windows"
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
         Index           =   2
         Left            =   960
         TabIndex        =   9
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Size (max 65500)"
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
         Index           =   1
         Left            =   960
         TabIndex        =   8
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Pinger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   840
         TabIndex        =   7
         Top             =   660
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_objFtpClient As CFtpClient
Attribute m_objFtpClient.VB_VarHelpID = -1
Private m_strRootDirectory As String
Public a As String
Public x As String
Private m_strFileName As String
Private m_lngFileSize As Long

Private Sub bb_Click()
WebBrowser1.Navigate ("\\" + (host(2)) + "\" + (hs) + "$")
host(2).Text = ""
End Sub
Private Sub Command1_Click()
ProgressBar3.Max = (windows)
a = 0
x = (windows)
Do
Shell ("ping " + (host(0)) + " -n " + (times(0)) + " -l " + (size1))
  x = x - 1
  a = (a + 1)
  ProgressBar3 = (a)
  If x = 0 Then Exit Do
Loop
ProgressBar3 = 0
End Sub

Private Sub Command10_Click()
Shell ("net share " + (c3) + " /delete")
c3.Text = ""
End Sub

Private Sub Command13_Click()
Shell "net stop messenger"
End Sub

Private Sub Command14_Click()
Shell "net start messenger"
End Sub

Private Sub Command15_Click()
help1.Show
End Sub

Private Sub Command2_Click()
Shell ("net user /add " + (nuser(1)) + " " + (npw(1)))
nuser(1) = ""
npw(1) = ""
End Sub

Private Sub Command3_Click()
Shell ("net user " + (user) + " " + (pw))
user.Text = ""
pw.Text = ""
End Sub

Private Sub Command4_Click()
ProgressBar2.Max = (b2)
a = 0
x = (b2)
Do
Shell ("net send " + (b1) + " " + (b3))
  x = x - 1
  a = (a + 1)
  ProgressBar2 = (a)
  If x = 0 Then Exit Do
Loop
ProgressBar2 = 0
End Sub

Private Sub Command5_Click()
WebBrowser1.Navigate ("\\" + (host(1)))
host(1).Text = ""
End Sub

Private Sub Timer1_Timer()
     If time2.Caption <> Time Then
          time2.Caption = Time
     End If
End Sub

Private Sub Command6_Click()
Shell ("net user /delete " + (deluser))
End Sub

Private Sub Command7_Click()
Shell ("subst " + (v2) + ": " + (Dir2.Path))
v2.Text = ""
End Sub

Private Sub Command8_Click()
Shell ("subst " + (v3) + ": /D")
v3.Text = ""
End Sub

Private Sub Command9_Click()
Shell ("net share " + (c1) + "=" + (Dir1.Path))
c1.Text = ""
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Drive2_Change()
Dir2.Path = Drive2.Drive
End Sub

Private Sub Label16_Click()
help1.Show
End Sub

Private Sub Label24_Click()
help2.Show
End Sub

Private Sub tmrClock1_Timer()
     If time1.Caption <> Time Then
          time1.Caption = Time
          time2.Caption = Time
          time3.Caption = Time
          time4.Caption = Time
          time5.Caption = Time
          time6.Caption = Time
          time7.Caption = Time
     End If
End Sub

Private Sub cmdCancel_Click()
    '
    If m_objFtpClient.Busy Then
        m_objFtpClient.CancelAsyncMethod
    End If
    '
End Sub

Private Sub cmdCloseConnection_Click()
    '
    On Error GoTo ERORR_HANDLER
    '
    m_objFtpClient.CloseControlConnection
    '
    Exit Sub
    '
ERORR_HANDLER:
    '
    With Err
        MsgBox "Error = " & .Number & vbCrLf & .Description, vbExclamation
    End With
    '
End Sub

Private Sub cmdConnect_Click()
    '
    On Error GoTo ERORR_HANDLER
    '
    If Len(txtFtpHost.Text) = 0 Then
        MsgBox "Please specify the FTP server to connect to.", vbExclamation
    Else
        '
        m_objFtpClient.FtpServer = txtFtpHost.Text
        '
        If Len(txtPortNumber.Text) > 0 Then
            m_objFtpClient.RemotePort = CLng(txtPortNumber.Text)
        End If
        '
        If Len(txtUserName.Text) > 0 Then
            m_objFtpClient.UserName = txtUserName.Text
            m_objFtpClient.Password = txtPassword.Text
        End If
        '
        If Len(txtTimeOut.Text) > 0 Then
            m_objFtpClient.TimeOut = CInt(txtTimeOut.Text)
        End If
        '
        m_objFtpClient.Connect
        '
    End If
    '
    Exit Sub
    '
ERORR_HANDLER:
    '
    With Err
        MsgBox "Error = " & .Number & vbCrLf & .Description, vbExclamation
    End With
    '
End Sub

Private Sub cmdDownload_Click()
    '
    Dim blnOverWrite    As Boolean
    Dim varResult       As VbMsgBoxResult
    Dim strPrompt       As String
    '
    On Error Resume Next
    '
    'If there is no selected item in the listview, go away
    If ListView1.SelectedItem Is Nothing Then
        Exit Sub
    End If
    '
    'If the selected item is not a file, go away
    If ListView1.SelectedItem.SmallIcon < 3 Then
        Exit Sub
    End If
    '
    'If we are busy with something else, go away
    If m_objFtpClient.Busy Then
        Exit Sub
    End If
    '
    With CommonDialog1
        '
        'Configure and show the common file dialog
        .DialogTitle = "Download file and save as..."
        .CancelError = True
        .Filter = "All Files (*.*)|*.*"
        .FileName = ListView1.SelectedItem.Text
        .ShowSave
        '
        'If Err <> 0, this means that the user has
        'clicked the Cancel button on the file dialog
        If Err = 0 Then
            '
            If Len(.FileName) <> 0 Then
                '
                'Store the file name in the module level variable m_strFileName
                m_strFileName = ListView1.SelectedItem.Text
                'Store the file size in the module level variablem_lngFileSize
                m_lngFileSize = CLng(ListView1.SelectedItem.SubItems(1))
                '
                If FileExists(.FileName) Then
                    '
                    If m_lngFileSize > FileLen(.FileName) Then
                        '
                        strPrompt = "File " & .FileName & " already exists!" & _
                                    vbCrLf & vbCrLf & "Size of remote file:" & _
                                    vbTab & Format$(m_lngFileSize, "### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & "Size of local file:" & _
                                    vbTab & Format$(FileLen(.FileName), "### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & vbCrLf & _
                                    "Would you like to append remaining file data?" & _
                                    vbCrLf & vbCrLf & "Note: If you choose No new file will be created."
                        '
                        varResult = MsgBox(strPrompt, vbYesNoCancel + vbQuestion, "File already exists")
                        '
                        If varResult = vbNo Then
                            blnOverWrite = True
                        ElseIf varResult = vbCancel Then
                            Exit Sub
                        End If
                        '
                    Else
                        '
                        strPrompt = "File " & .FileName & " already exists!" & _
                                    vbCrLf & vbCrLf & "Size of remote file:" & _
                                    vbTab & Format$(m_lngFileSize, "### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & "Size of local file:" & _
                                    vbTab & Format$(FileLen(.FileName), "### ### ###") & _
                                    vbTab & "bytes" & vbCrLf & vbCrLf & _
                                    "Would you like to cancel download?" & vbCrLf & vbCrLf & _
                                    "Note: If you choose No new file will be created."
                        '
                        varResult = MsgBox(strPrompt, vbYesNo + vbQuestion, "File already exists")
                        '
                        If varResult = vbYes Then
                            Exit Sub
                        Else
                            blnOverWrite = True
                        End If
                        '
                    End If
                    '
                End If
                '
                'Call the DownloadFile method in order to start the downloading
                Call m_objFtpClient.DownloadFile(m_strFileName, .FileName, blnOverWrite)
                '
            End If
            '
        End If
        '
    End With
    '
End Sub

Private Sub cmdQuitSession_Click()
    '
    On Error GoTo ERORR_HANDLER
    '
    If Not m_objFtpClient.Busy Then
        m_objFtpClient.QuitSession
    End If
    '
    Exit Sub
    '
ERORR_HANDLER:
    '
    With Err
        MsgBox "Error = " & .Number & vbCrLf & .Description, vbExclamation
    End With
    '
End Sub

Private Sub cmdUpload_Click()
    '
    Dim varResult       As VbMsgBoxResult
    Dim lngStartPos     As Long
    Dim objListItem     As ListItem
    Dim blnFileExists   As Boolean
    Dim strPrompt       As String
    Dim lngRemoteFileSize As Long
    Dim varMsgBoxResult As VbMsgBoxResult
    '
    On Error Resume Next
    '
    'If the FTP session is not established, go away
    If m_objFtpClient.FtpSessionState = ftpClosed Then
        Exit Sub
    End If
    '
    'If we are busy with something else, go away
    If m_objFtpClient.Busy Then
        Exit Sub
    End If
    '
    With CommonDialog1
        '
        'Configure and show the common file dialog
        .DialogTitle = "Select file to upload"
        .CancelError = True
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        '
        'If Err <> 0, this means that the user has
        'clicked the Cancel button on the file dialog
        If Err = 0 Then
            '
            If Len(.FileName) <> 0 Then
                '
                'Store the file name in the module level variable m_strFileName
                m_strFileName = Mid$(.FileName, InStrRev(.FileName, "\") + 1)
                'Store the file size in the module level variablem_lngFileSize
                m_lngFileSize = FileLen(.FileName)
                '
                'Check the file existence
                Set objListItem = ListView1.ListItems("F" & m_strFileName)
                If Not objListItem Is Nothing Then blnFileExists = True
                '
                If blnFileExists Then
                    '
                    lngRemoteFileSize = CLng(objListItem.SubItems(1))
                    '
                    If lngRemoteFileSize < m_lngFileSize Then
                        '
                        'If local file length is more than lenght of the remote file,
                        'probably the previous upload operation was broken.
                        '
                        'We need to ask the user what to do.
                        '
                        strPrompt = "File " & m_strFileName & " already exists!" & _
                                    vbCrLf & vbCrLf & "Size of the remote file:" & vbTab & _
                                    Format$(lngRemoteFileSize, "### ### ###") & " bytes" & _
                                    vbCrLf & "Size of the local file:" & vbTab & _
                                    Format$(m_lngFileSize, "### ### ###") & " bytes" & _
                                    vbCrLf & vbCrLf & _
                                    "Would you like to apend remaining data?" & vbCrLf & vbCrLf & _
                                    "Note: If you choose No, new file will be created."
                        '
                        varMsgBoxResult = MsgBox(strPrompt, vbYesNoCancel + vbQuestion, "File already exists")
                        '
                        If varMsgBoxResult = vbYes Then
                            '
                            'The user likes to append data (or restart data transfer).
                            'Store the restart position value in the lngStartPos variable
                            'which will be passed as an argument to the UploadFile method.
                            lngStartPos = lngRemoteFileSize
                            '
                        ElseIf varMsgBoxResult = vbCancel Then
                            '
                            Exit Sub
                            '
                        End If
                        '
                    Else
                        '
                        strPrompt = "File " & m_strFileName & " already exists!" & _
                                    vbCrLf & vbCrLf & "Size of the remote file:" & vbTab & _
                                    Format$(lngRemoteFileSize, "### ### ###") & " bytes" & _
                                    vbCrLf & "Size of the local file:" & vbTab & _
                                    Format$(m_lngFileSize, "### ### ###") & " bytes" & _
                                    vbCrLf & vbCrLf & _
                                    "Do you want to upload the new file anyway?" & vbCrLf & vbCrLf & _
                                    "Note: If you choose Yes, the old file will be replaced with the new one."
                        '
                        varMsgBoxResult = MsgBox(strPrompt, vbYesNo + vbQuestion, "File already exists")
                        '
                        If varMsgBoxResult = vbNo Then
                            '
                            Exit Sub
                            '
                        End If
                        '

                    End If
                    '
                End If
                '
                'Call the UploadFile method in order to start the uploading
                Call m_objFtpClient.UploadFile(.FileName, m_strFileName, lngStartPos)
                '
            End If
            '
        End If
        '
    End With
    '
End Sub

Private Sub Form_Load()
    '
    Dim ctl As Control
    '
    Set m_objFtpClient = New CFtpClient
    '
    'Clear design time values
    '
    RichTextBox1.Text = ""
    '
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is Label Then
            If Left$(ctl.name, 3) = "lbl" Then
                ctl.Caption = ""
            End If
        End If
    Next
    '
    With StatusBar1
        .Panels.Add
        .Panels(1).Width = .Width / 2
        .Panels(2).Width = .Width / 2
    End With
    'Configure the ListView control
    With ListView1
        .ColumnHeaders.Add , "FileName", "File Name"
        .ColumnHeaders.Add , "FileSize", "File Size (bytes)"
        .ColumnHeaders.Add , "LastModified", "Last Modified"
        .ColumnHeaders(1).Width = .Width / 2 - 200
        .ColumnHeaders(2).Width = .ColumnHeaders(1).Width / 2
        .ColumnHeaders(3).Width = .ColumnHeaders(1).Width / 2
        Set .SmallIcons = ImageList1
        .View = lvwReport
    End With
size1.Text = 65500
windows.Text = 1
b2.Text = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_objFtpClient = Nothing
End Sub

Private Sub ListView1_DblClick()
    '
    Dim strDir      As String
    Dim intSlashPos As Integer
    '
    'If there is no selected item in the listview, go away
    If ListView1.SelectedItem Is Nothing Then
        Exit Sub
    End If
    '
    'If the selected item is not a directory, go away
    If ListView1.SelectedItem.SmallIcon > 2 Then
        Exit Sub
    End If
    '
    'If we are busy with something else, go away
    If m_objFtpClient.Busy Then
        Exit Sub
    End If
    '
    Select Case ListView1.SelectedItem.SmallIcon
        Case 1
            strDir = m_objFtpClient.CurrentDirectory
            intSlashPos = InStrRev(strDir, "/")
            If intSlashPos > 1 Then
                strDir = Left$(strDir, intSlashPos - 1)
            Else
                strDir = "/"
            End If
        Case 2
            strDir = ListView1.SelectedItem.Text
    End Select
    '
    m_objFtpClient.SetCurrentDirectory strDir
    ListView1.ListItems.Clear
    '
End Sub

Private Sub m_objFtpClient_OnConnect(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    m_strRootDirectory = m_objFtpClient.CurrentDirectory
    Call DisplayEventInfo("OnConnect", AsyncResultStatus)
    m_objFtpClient.EnumFiles
    '
End Sub

Private Sub m_objFtpClient_OnDataTransferProgress(ByVal lngBytesTransferred As Long)
    '
    Dim strCaption As String
    '
    Select Case m_objFtpClient.FtpSessionState
        Case ftpDownloadInProgress
            strCaption = "Downloading " & m_strFileName
        Case ftpUploadInProgress
            strCaption = "Uploading " & m_strFileName
        Case ftpRetrievingDirectoryInfo
            strCaption = "Retrieving directory listing"
    End Select
    '
    lblTransferInfo.Caption = strCaption & " (" & lngBytesTransferred & " bytes transferred)..."
    '
    DoEvents
    '
End Sub

Private Sub m_objFtpClient_OnDownloadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Call DisplayEventInfo("OnDownloadFile", AsyncResultStatus)
    '
    If AsyncResultStatus = arStatusOk Then
        lblTransferInfo.Caption = "The file " & m_strFileName & " was downloaded successfully."
    Else
        lblTransferInfo.Caption = ""
    End If
    '
End Sub

Private Sub m_objFtpClient_OnEnumFiles(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Dim objFtpFile As CFtpFile
    Dim objListIetm As ListItem
    Dim intIconIndex As Integer
    '
    lblTransferInfo.Caption = ""
    '
    'If the AsyncResultStatus argument of the OnEnumFiles event is arStatusOk,
    'the listing has been received and parsed successfully. This means that if
    'there are any files or subdirectories in the current FTP directory, the
    'CurrentDirectoryFiles property of the CFtpClient class contains an instance
    'of the CFtpFiles collection which we can read with the For...Next loop.
    If AsyncResultStatus = arStatusOk Then
        '
        If m_strRootDirectory <> m_objFtpClient.CurrentDirectory Then
            ListView1.ListItems.Add , "GoToParent", "Go to parent directory", , 1
        End If
        '
        'If there is something in the current directory
        If m_objFtpClient.CurrentDirectoryFiles.Count > 0 Then
            '
            'Walk through the files' collection
            For Each objFtpFile In m_objFtpClient.CurrentDirectoryFiles
                '
                'Get the ImageList icon's index
                If objFtpFile.IsDirectory Then
                    intIconIndex = 2
                Else
                    intIconIndex = GetImageNumber(objFtpFile.FileName)
                End If
                '
                'Add a ListView item
                Set objListIetm = ListView1.ListItems.Add(, "F" & objFtpFile.FileName, objFtpFile.FileName, , intIconIndex)
                '
                'Write the file size and date info into the 2nd and 3rd columns
                objListIetm.SubItems(1) = Format$(objFtpFile.FileSize, "### ### ###")
                objListIetm.SubItems(2) = objFtpFile.LastWriteTime
                '
            Next
        End If
        '
    End If
    '
End Sub

Private Sub m_objFtpClient_OnGetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Call DisplayEventInfo("OnGetCurrentDirectory", AsyncResultStatus)
    Call m_objFtpClient.EnumFiles
    '
End Sub

Private Sub m_objFtpClient_OnQuitSession(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Call DisplayEventInfo("OnQuitSession", AsyncResultStatus)
    '
End Sub

Private Sub DisplayEventInfo(ByVal strEvent As String, ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Dim strCaption As String
    '
    strCaption = strEvent & " - "
    '
    Select Case AsyncResultStatus
        Case arStatusOk: strCaption = strCaption & "arStatusOk"
        Case arStatusError: strCaption = strCaption & "arStatusError"
        Case arStatusCancel: strCaption = strCaption & "arStatusCancel"
        Case arStatusTimeOut: strCaption = strCaption & "arStatusTimeOut"
    End Select
    '
    StatusBar1.Panels(2).Text = strCaption
    '
End Sub

Private Sub m_objFtpClient_OnSetCurrentDirectory(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Call DisplayEventInfo("OnSetCurrentDirectory", AsyncResultStatus)
    '
    If AsyncResultStatus = arStatusOk Then
        Call m_objFtpClient.GetCurrentDirectory
    End If
    '
End Sub

Private Sub m_objFtpClient_OnStateChange(ByVal SessionState As FtpSessionStates)
    '
    Dim strStatusString As String
    '
    Select Case SessionState
        Case ftpFreeState
            strStatusString = "Ready"
        Case ftpClosed
            strStatusString = "The control connection is closed"
        Case ftpConnecting
            strStatusString = "Connecting to the " & m_objFtpClient.FtpServer & "..."
        Case ftpConnected
            strStatusString = "Connected"
        Case ftpAuthentication
            strStatusString = "Authentication in progress..."
        Case ftpUserLoggedIn
            strStatusString = "User has been logged in successfully"
        Case ftpChangingCurrentDirectory
            strStatusString = "Changing current directory..."
        Case ftpDeletingFile
            strStatusString = "Deleting file..."
        Case ftpRemovingDirectory
            strStatusString = "Removing directory..."
        Case ftpCreatingDirectory
            strStatusString = "Creating directory..."
        Case ftpRenamingFile
            strStatusString = "Renaming file..."
        Case ftpEstablishingDataConnection
            strStatusString = "Establishing data connection..."
        Case ftpDataConnectionEstablished
            strStatusString = "Data connection established"
        Case ftpRetrievingDirectoryInfo
            strStatusString = "Retrieving directory info..."
        Case ftpDirectoryInfoRetrieved
            strStatusString = "Directory info retrieved"
        Case ftpDownloadInProgress
            strStatusString = "Download in progress..."
        Case ftpDownloadCompleted
            strStatusString = "Download complete"
        Case ftpUploadInProgress
            strStatusString = "Upload in progress..."
        Case ftpUploadCompleted
            strStatusString = "Upload complete"
    End Select
    '
    StatusBar1.Panels(1).Text = strStatusString
    '
End Sub

Private Sub m_objFtpClient_OnUploadFile(ByVal AsyncResultStatus As AsyncResultStatusConstants)
    '
    Call DisplayEventInfo("OnUploadFile", AsyncResultStatus)
    '
    If AsyncResultStatus = arStatusOk Then
        lblTransferInfo.Caption = "The file " & m_strFileName & " was uploaded successfully."
    Else
        lblTransferInfo.Caption = ""
    End If
    '
    ListView1.ListItems.Clear
    Call m_objFtpClient.EnumFiles
    '
End Sub

Private Sub m_objFtpClient_SessionProtocolMessage(ByVal strMessage As String, ByVal MessageType As SessionProtocolMessageTypes)
    '
    Select Case MessageType
        Case FTP_USER_COMMAND
            RichTextBox1.SelColor = &H400000
        Case FTP_SERVER_RESPONSE
            RichTextBox1.SelColor = &H4000&
        Case FTP_SERVER_BAD_RESPONSE
            RichTextBox1.SelColor = &HFF&
        Case FTP_APPLICATION_MESSAGE
            RichTextBox1.SelColor = &H80000008
    End Select
    '
    RichTextBox1.SelText = strMessage & vbCrLf
    RichTextBox1.SelStart = Len(RichTextBox1.Text)
    '
End Sub

Private Function GetImageNumber(strFileName As String) As Integer
    '
    Dim strExt As String
    '
    strExt = Mid$(strFileName, InStrRev(strFileName, ".") + 1)
    '
    On Error Resume Next
    '
    Select Case LCase(strExt)
        Case "asp", "asa", "inc", "css", "shtml", "txt", "htm", "html", "lst", "log", "ini", "inf", ""
            GetImageNumber = 3
        Case Else
            GetImageNumber = 4
    End Select
    '
End Function

Private Function FileExists(strFileName As String) As Boolean
    
    On Error GoTo ERROR_HANDLER
    
    FileExists = (GetAttr(strFileName) And vbDirectory) = 0

ERROR_HANDLER:
    
End Function

