VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Balls!!!!"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   3720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer run 
      Interval        =   1
      Left            =   720
      Top             =   2760
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   480
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Hrayr Artunyan
' June 12, 2004
' hrayr@artunyan.com (for any comments or questions)
' www.artunyan.com
' This is a small fun program with absolutely no practical use. It is written for the sole
' purpose of demonstration. Have fun and please let me know what you think!

' Try resizing the window or moving the window around; what happens?

Option Explicit

Dim ball() As ball          ' my projectile type
Dim wrld_width As Single    ' world width (picture box width)
Dim wrld_height As Single   ' world height (picture box height)
Dim ball_width As Single    ' width of the ball (picture box DrawWidth)

Private Sub Form_Load()
    Dim num_balls As Single
    Dim br_min As Single        ' declare minimum ball radius
    Dim br_max As Single        ' declare maximum ball radius
    
    br_min = 5                  ' set minimum ball radius
    br_max = 10                 ' set maximum ball radius
    ' setup picturebox.
    picBox.FillStyle = vbFSSolid
    picBox.DrawMode = 10                ' not xor pen
    picBox.DrawWidth = 2                ' width of the circles
    ball_width = picBox.DrawWidth
    num_balls = 5                      ' number of balls in the box.
    ReDim ball(num_balls - 1)
    
    Dim i As Integer
    ' initialize balls to random positions, velocities, size and color
    For i = 0 To UBound(ball)
        Randomize
        ball(i).r = br_min + Rnd() * (br_max - br_min)                      ' set rand radius of ball
        ball(i).x = ball(i).r + Rnd() * (wrld_width - ball(i).r * 2) ' set rand x position
        ball(i).y = Rnd * (wrld_height - ball(i).r)                  ' set rand y position
        ball(i).xv = (0.5 - Rnd()) * 5                                      ' set rand x_vel
        ball(i).yv = (0.5 - Rnd()) * 5                                      ' set rand y_vel
        ball(i).c = RGB(Rnd * 255, Rnd * 255, Rnd * 255)                    ' set random color
    Next i
    
End Sub

Private Sub Form_Resize()
    ' resize picture box to window size
    picBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    wrld_width = picBox.ScaleWidth
    wrld_height = picBox.ScaleHeight
    DoEvents
End Sub

Private Sub run_Timer()
    Dim j As Single
    Static i As Integer
    
    Static wrld_left As Single     ' form left position
    Static wrld_top As Single      ' form top position
    Static wrld_wdth As Single     ' width of window
    Static wrld_hght As Single     ' height of window
    
    If i < 2 Then
        wrld_left = 0
        wrld_top = 0
        wrld_wdth = 0
        wrld_hght = 0
        i = i + 1
    End If
    
    ' how much did the window move or resize? We take the position of the window into account
    ' when calculating position of the balls. Look for these variables in the following code
    wrld_left = (wrld_left - Me.Left) / 15
    wrld_top = (wrld_top - Me.Top) / 15
    wrld_wdth = (wrld_wdth - wrld_width) / 15
    wrld_hght = (wrld_hght - wrld_height) / 15
    
    ' print my name on the top left corner ( you can take this out if you want )
    picBox.CurrentX = 3
    picBox.CurrentY = 3
    picBox.ForeColor = vbBlue
    picBox.Print "by Hrayr Artunyan"
    
    For j = 0 To UBound(ball)
        ' erase previous balls. you can also use picBox.cls before the for loop but it's slower like that
        If i > 1 Then
            picBox.FillColor = ball(j).c
            picBox.ForeColor = ball(j).c
            picBox.Circle (ball(j).x, ball(j).y), ball(j).r
        End If

        ' kinematics ( laws of motion )
        ball(j).x = ball(j).x + ball(j).xv + wrld_left       ' remember x = x + x_vel
                                                    ' there is no air friction so x_vel is always constant in mid air
                                                    ' there is friction on the floor and it's calculated in "hit floor"
        ball(j).xv = ball(j).xv
        ball(j).y = ball(j).y + ball(j).yv + wrld_top        ' y = y + y_vel
        ball(j).yv = ball(j).yv + 0.1              ' y_vel = y_vel + accel*t       t is 1 interval so we don't write it
        
        ' ***wall collision detection***
        ' I left the ceiling out
        If (ball(j).y - ball(j).r - ball_width / 2) < 0 Then
            ball(j).y = ball(j).r + ball_width / 2
            ball(j).yv = -ball(j).yv - 0.2 + wrld_top / 15
        End If
        
        ' hit floor
        If (ball(j).y + ball(j).r + ball_width / 2) >= wrld_height Then
            ball(j).y = wrld_height - ball_width / 2 - ball(j).r    ' bring ball back from underground
            ball(j).yv = -ball(j).yv + 1 + wrld_top / 15                       ' change direction of velocity and loose some energy
            ball(j).xv = ball(j).xv - ball(j).xv * 0.01 - wrld_left / 60       ' this is where friction comes in
        End If
        
        ' hit right wall
        If (ball(j).x + ball(j).r + ball_width / 2) >= wrld_width Then
            ball(j).x = wrld_width - ball_width / 2 - ball(j).r     ' position ball so it doesn't go through the wall
            ball(j).xv = -ball(j).xv + 1 + wrld_left / 15           ' change velocity direction and loose energy
        End If
        
        'hit left wall
        If (ball(j).x - ball(j).r - ball_width / 2) <= 0 Then
            ball(j).x = ball(j).r + ball_width / 2                  ' position ball so it doesn't go through the wall
            ball(j).xv = -ball(j).xv - 1 + wrld_left / 15        ' change velocity direction and loose energy
        End If
        ' ***end collision detection***
        
        ' draw ball
        picBox.FillColor = ball(j).c
        picBox.ForeColor = ball(j).c
        picBox.Circle (ball(j).x, ball(j).y), ball(j).r
    Next j
    
    ' set world variables ( position and size of windows )
    wrld_top = Me.Top
    wrld_left = Me.Left
    wrld_wdth = wrld_width
    wrld_hght = wrld_height
    DoEvents
End Sub
