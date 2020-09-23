VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form startup 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab help 
      Height          =   4455
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   7858
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Help"
      TabPicture(0)   =   "startup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         Height          =   195
         Left            =   1320
         TabIndex        =   15
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"startup.frx":001C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   3000
         Width           =   3615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Other Things:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   $"startup.frx":00CB
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1815
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "What happens:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Arrow Keys"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "Movement:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      Height          =   495
      Left            =   2880
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   " Help"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose and press ENTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   3600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Darwin Yu"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Created by:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Racer "
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tank"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Index           =   0
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   1080
      Picture         =   "startup.frx":01A2
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
help.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
If Shape1.Top > Label9.Top Then
Shape1.Top = Label8.Top
Else
Shape1.Top = Shape1.Top + 480
End If
End If
If KeyCode = vbKeyUp Then
If Shape1.Top < Label9.Top Then
Shape1.Top = 1920
Else
Shape1.Top = Shape1.Top - 480
End If
End If
If KeyCode = vbKeyReturn Then
If Shape1.Top = Label8.Top Then
startup.Hide
Load track1
track1.Show
End If
If Shape1.Top = Label10.Top Then
End
End If
If Shape1.Top = Label9.Top Then
help.Visible = True
End If
End If
End Sub

