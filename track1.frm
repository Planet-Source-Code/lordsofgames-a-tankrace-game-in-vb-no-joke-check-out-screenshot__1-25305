VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form track1 
   BackColor       =   &H00404040&
   Caption         =   "Tank Race v.1.1."
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   Icon            =   "track1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "track1.frx":164A
   ScaleHeight     =   7650
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer badguy2crash 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   1920
   End
   Begin VB.Timer badguy1crash 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   1920
   End
   Begin VB.Timer collide 
      Interval        =   80
      Left            =   3360
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      Height          =   195
      Left            =   8400
      TabIndex        =   14
      Top             =   7680
      Width           =   255
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":42D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":592C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":6F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":85E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer goodguycrash 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   1920
   End
   Begin VB.Timer badguyspeed 
      Interval        =   1400
      Left            =   3000
      Top             =   3480
   End
   Begin VB.Timer badguy2controllerb 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2640
      Top             =   4080
   End
   Begin VB.Timer badguy2controller 
      Interval        =   50
      Left            =   2280
      Top             =   4080
   End
   Begin VB.Timer badguy1controllerb 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2760
      Top             =   3000
   End
   Begin VB.Timer badguy1controller 
      Interval        =   50
      Left            =   2280
      Top             =   3000
   End
   Begin VB.PictureBox bad2right 
      Height          =   735
      Left            =   6120
      Picture         =   "track1.frx":9C40
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad2left 
      Height          =   735
      Left            =   6000
      Picture         =   "track1.frx":B28A
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   12
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad2down 
      Height          =   735
      Left            =   5880
      Picture         =   "track1.frx":C8D4
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad2up 
      Height          =   735
      Left            =   5760
      Picture         =   "track1.frx":DF1E
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad1right 
      Height          =   735
      Left            =   3120
      Picture         =   "track1.frx":F568
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad1left 
      Height          =   735
      Left            =   3000
      Picture         =   "track1.frx":10BB2
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad1down 
      Height          =   735
      Left            =   2880
      Picture         =   "track1.frx":121FC
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox bad1up 
      Height          =   735
      Left            =   2760
      Picture         =   "track1.frx":13846
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox goodright 
      Height          =   735
      Left            =   720
      Picture         =   "track1.frx":14E90
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox goodleft 
      Height          =   735
      Left            =   600
      Picture         =   "track1.frx":164DA
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox gooddown 
      Height          =   735
      Left            =   480
      Picture         =   "track1.frx":17B24
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox goodup 
      Height          =   735
      Left            =   360
      Picture         =   "track1.frx":1916E
      ScaleHeight     =   675
      ScaleWidth      =   795
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4080
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":1A7B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":1BE14
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":1D470
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":1EACC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5520
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":20128
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":21784
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":22DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "track1.frx":2443C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gas Station -->"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5640
      TabIndex        =   15
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Image badguy2 
      Height          =   960
      Left            =   6120
      Picture         =   "track1.frx":25A98
      Top             =   6240
      Width           =   960
   End
   Begin VB.Image badguy1 
      Height          =   960
      Left            =   5280
      Picture         =   "track1.frx":270E2
      Top             =   6240
      Width           =   960
   End
   Begin VB.Image goodguy 
      Height          =   960
      Left            =   6960
      Picture         =   "track1.frx":2872C
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Start"
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
      Left            =   6120
      TabIndex        =   1
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label done 
      BackStyle       =   0  'Transparent
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Image grass1 
      Height          =   5565
      Left            =   8040
      Picture         =   "track1.frx":29D76
      Top             =   2400
      Width           =   420
   End
   Begin VB.Image grass3 
      Height          =   5565
      Left            =   0
      Picture         =   "track1.frx":2C9FC
      Top             =   2160
      Width           =   420
   End
   Begin VB.Image curb1 
      Height          =   5835
      Left            =   4080
      Picture         =   "track1.frx":2F682
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Image curb2 
      Height          =   5835
      Left            =   3000
      Picture         =   "track1.frx":367DC
      Top             =   3720
      Width           =   1065
   End
   Begin VB.Image grass2 
      Height          =   1770
      Left            =   -120
      Picture         =   "track1.frx":3D936
      Top             =   0
      Width           =   9690
   End
End
Attribute VB_Name = "track1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tank Racer in VB
'Created and programmed by Darwin Yu
'Well, this is the racing game I made.
'Sorry, if it's not much to you,
'I just made it for enjoyment.
'And yes, there are a lotta timers,
'I don't have another way of doing this,
'I wanted things to happen with different intervals,
'Anyways, if you like it, please vote.
'If you don't like it, please don't vote poorly.
'Thanks!
'You can make more tracks(which require more forms) if
'you want to. But it'll take lots of modification.
'Please do not take any graphics.
'Enjoy!
Private Function Collision() As Boolean ' Collision from the player's point of view.
    If Not ((goodguy.Left > badguy1.Left + 360) Or (goodguy.Left + 240 < badguy1.Left)) Then
        Collision = Not ((goodguy.Top > badguy1.Top + 360) Or (goodguy.Top + 360 < badguy1.Top))
    End If
End Function
Private Function Collision2() As Boolean ' Collision from player3's point of view.
    If Not ((badguy2.Left > badguy1.Left + 360) Or (badguy2.Left + 240 < badguy1.Left)) Then
        Collision2 = Not ((badguy2.Top > badguy1.Top + 360) Or (badguy2.Top + 360 < badguy1.Top))
    End If
End Function
Private Function Collision3() As Boolean ' Collision from player 2's point of view.
    If Not ((badguy1.Left > badguy2.Left + 360) Or (badguy1.Left + 240 < badguy2.Left)) Then
        Collision3 = Not ((badguy1.Top > badguy2.Top + 360) Or (badguy1.Top + 360 < badguy2.Top))
    End If
End Function
Private Function Collision4() As Boolean ' Collision from the player's point of view.
    If Not ((goodguy.Left > badguy2.Left + 360) Or (goodguy.Left + 240 < badguy2.Left)) Then
        Collision4 = Not ((goodguy.Top > badguy2.Top + 360) Or (goodguy.Top + 360 < badguy2.Top))
    End If
End Function

Private Sub badguy1controller_Timer()
If badguy1.Top < (curb1.Top - 1000) Then
badguy1.Picture = bad1left.Picture
badguy1.Left = badguy1.Left - 40
Else
badguy1.Top = badguy1.Top - 40
End If
If badguy1.Left < (curb2.Left - 1000) Then
badguy1controllerb.Enabled = True
badguy1controller.Enabled = False
End If
End Sub

Private Sub badguy1controllerb_Timer()
If badguy1.Left < (curb2.Left - 1000) Then
badguy1.Picture = bad1down.Picture
badguy1.Top = badguy1.Top + 40
End If
End Sub

Private Sub badguy1crash_Timer()
On Error Resume Next
Static idx As Integer, Counter As Integer
If Counter > 40 Then 'If the car spinned 40 times,
badguy1crash.Enabled = False
Else
idx = (idx + 1) Mod 5    '5 Frames to work with and a lotta spinning
If idx = 0 Then idx = 1 'to do
badguy1.Picture = ImageList2.ListImages(idx).Picture
Counter = Counter + 1
End If
End Sub

Private Sub badguy2controller_Timer()
If badguy2.Top < (curb1.Top - 1200) Then
badguy2.Picture = bad2left.Picture
badguy2.Left = badguy2.Left - 40
Else
badguy2.Top = badguy2.Top - 40
End If
If badguy2.Left < (curb2.Left - 1200) Then
badguy2controllerb.Enabled = True
badguy2controller.Enabled = False
End If
End Sub

Private Sub badguy2controllerb_Timer()
If badguy2.Left < (curb2.Left - 1200) Then
badguy2.Picture = bad2down.Picture
badguy2.Top = badguy2.Top + 40
End If
End Sub

Private Sub badguy2crash_Timer()
On Error Resume Next
Static idx As Integer, Counter As Integer
If Counter > 40 Then 'If the car spinned 40 times,
badguy2crash.Enabled = False
Else
idx = (idx + 1) Mod 5    '5 Frames to work with and a lotta spinning
If idx = 0 Then idx = 1 'to do
badguy2.Picture = ImageList3.ListImages(idx).Picture
Counter = Counter + 1
End If
End Sub

Private Sub badguyspeed_Timer()
badguy1controller.Interval = Int((120 * Rnd) + 1)
badguy2controller.Interval = Int((120 * Rnd) + 1)
badguy1controllerb.Interval = Int((60 * Rnd) + 1)
badguy2controllerb.Interval = Int((58 * Rnd) + 1)

End Sub

Private Sub collide_Timer()
'sorry for this new timer, but
'I just couldn't find a place
'to detect the collision
'and crashes
If Collision = True Or Collision4 = True Then
goodguycrash.Enabled = True
Command1.Enabled = True
End If
If Collision2 = True Then
badguy1crash.Enabled = True
End If
If Collision3 = True Then
badguy2crash.Enabled = True
End If
'If goodguy.Top > curb1.Top And goodguy.Left < curb1.Left Then
'goodguycrash.Enabled = True
'MsgBox "Never Cheat like that!" I can't get the collisoin on here.
'Unload Me
'End
'End If
'If goodguy.Left < (curb2.Left + 700) Then
'goodguycrash.Enabled = True
'MsgBox "Never Cheat like that!"
'Unload Me
'End
'End If
If goodguy.Top > done.Top And goodguy.Top > badguy1.Top And badguy1.Top > badguy2.Top Then
MsgBox ("First Place: You" & vbNewLine & "Second Place: Blue Tank" & vbNewLine & "Third Place: Green Tank")
Unload Me
End
End If
If goodguy.Top > done.Top And goodguy.Top > badguy2.Top And badguy2.Top > badguy1.Top Then
MsgBox ("First Place: You" & vbNewLine & "Second Place: Green Tank" & vbNewLine & "Third Place: Blue Tank")
Unload Me
End
End If
If badguy1.Top > done.Top And badguy1.Top > goodguy.Top And goodguy.Top > badguy2.Top Then
MsgBox ("First Place: Blue Tank" & vbNewLine & "Second Place: You" & vbNewLine & "Third Place: Green Tank")
Unload Me
End
End If
If badguy2.Top > done.Top And badguy2.Top > goodguy.Top And goodguy.Top > badguy1.Top Then
MsgBox ("First Place: Green Tank" & vbNewLine & "Second Place: You" & vbNewLine & "Third Place: Blue Tank")
Unload Me
End
End If
If badguy2.Top > done.Top And badguy2.Top > badguy1.Top And badguy1.Top > goodguy.Top Then
MsgBox ("First Place: Green Tank" & vbNewLine & "Second Place: Blue Tank" & vbNewLine & "Third Place: You")
Unload Me
End
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
goodguy.Picture = goodup.Picture
goodguy.Top = goodguy.Top - Int((120 * Rnd) + 1)
End If
If KeyCode = vbKeyDown Then
goodguy.Picture = gooddown.Picture
goodguy.Top = goodguy.Top + Int((120 * Rnd) + 1)
End If
If KeyCode = vbKeyLeft Then
goodguy.Picture = goodleft.Picture
goodguy.Left = goodguy.Left - Int((120 * Rnd) + 1)
End If
If KeyCode = vbKeyRight Then
goodguy.Picture = goodright.Picture
goodguy.Left = goodguy.Left + Int((120 * Rnd) + 1)
End If
End Sub

Private Sub goodguycrash_Timer()
On Error Resume Next
Static idx As Integer, Counter As Integer
If Counter > 40 Then 'If the car spinned 40 times,
Command1.Enabled = False
goodguycrash.Enabled = False
Else
idx = (idx + 1) Mod 5    '5 Frames to work with and a lotta spinning
If idx = 0 Then idx = 1 'to do
goodguy.Picture = ImageList1.ListImages(idx).Picture
Counter = Counter + 1
End If
End Sub

Private Sub Timer1_Timer()

End Sub
