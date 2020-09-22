VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFF00&
   Caption         =   "Form1"
   ClientHeight    =   11145
   ClientLeft      =   1455
   ClientTop       =   315
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   ScaleHeight     =   743
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   925
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPlane 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   6000
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   1800
   End
   Begin VB.Timer tmrBG 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   9120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    BG_X = 0
    BitBlt frmMain.hdc, 0, 0, Me.Width, Me.Height, BG, BG_X, 0, vbSrcCopy
End Sub

Private Sub Form_DblClick()
    Plane_Y = 1000 'Plane start being drawn near the ground
    
    PlanePicX = 120 '
    PlanePicY = 80  'This tells the program to draw the plane straight right
    
    picPlane.Left = frmMain.ScaleWidth / 2 - picPlane.Width / 2
    
    TakeOff = True
    
    MaxVel = 5
    UpDown = 0
    Run_Game
End Sub


Private Sub Form_Load()
    'This is inside a while loop because sometimes the computer gets an
    'error while trying to load pictures. This just makes sure it keeps
    'trying until if finally gets is
    While BG <= 0
        BG = GenerateDC(App.Path + "\BGBigHSmall.bmp") 'BG now points at the spot in memory where the picture is placed
    Wend
    
    While PlaneMask <= 0
        PlaneMask = GenerateDC(App.Path + "\PlaneMask.bmp")
    Wend
    
    While PlaneSprite <= 0
        PlaneSprite = GenerateDC(App.Path + "\PlaneSprite.bmp")
    Wend
    
'    While TreeMask <= 0                                        'These are for trees, I was going to blt them on top of the
'        TreeMask = GenerateDC(App.Path + "\TreesMask.bmp")     'bckgrnd so when/if the plane crashed because of the trees
'    Wend                                                       'you would get the effect of the plain crashing behind the
'                                                               'trees but it took up too much processor speed so I took it
'    While TreeSprite <= 0                                      'out. If you have a fast enough processor, you might want to
'        TreeSprite = GenerateDC(App.Path + "\TreesSprite.bmp") 'include this, just remember to calibrate it's positions
'    Wend                                                       'using TreeX
    
    Plane_Y = 1000 'Plane start being drawn near the ground
    
    PlanePicX = 120 '
    PlanePicY = 80  'This tells the program to draw the plane straight right
    
    picPlane.Left = frmMain.ScaleWidth / 2 - picPlane.Width / 2
    
    TakeOff = True
    
    MaxVel = 5
    
    UpDown = 0
    
    TreeX = 3000
End Sub

Public Function Run_Game()
    'GetAsyncKeyState is a function that checks the state of a key, meaning it checks whether it is up or
    'down. If it is up(not pressed), the function returns 0, if it is down, it will return -32767 the first
    'time it checks the key. If the key is being held and the function keeps checking, it will incrememnt
    '-32767 by -1, it's rather confusing but just remember: If you want to check if it's simply being pushed
    'or not, use "<> 0" If you're checking for the first time it's pushed though, use "= -32767" Another way
    'to determine which to use is if the computer is checking it too fast, use the -32767, that should slow
    'it down and vice versa if it is too slow
    Do
        'Me.Refresh
        If TakeOff = True And Abs(LeftRight) > 4 Then TakeOff = False
        PlaneChange = False
        If UpDown >= 1.5 Then       '
            PlanePicY = 160         ' This just checks to see how fast the plane is moving up and down, if
        ElseIf UpDown <= -1.5 Then  ' is is moving fast in either direction, it makes it so the plane is
            PlanePicY = 0           ' facing that direction. In other words, if the plane is going up steeply
        Else                        ' this will make it so it is facing upwards, instead of straight forward
            PlanePicY = 80           '
        End If                      '
        If TurningRight = True Then
            '******DO TURNING ANIMATION HERE*******
            If LeftRight < 3 Then
                LeftRight = LeftRight + 0.2
            Else
                TurningRight = False
            End If
            If Int(LeftRight) = 0 Then PlanePicX = 120
        ElseIf TurningLeft = True Then
            '******DO TURNING ANIMATION HERE*******
            If LeftRight > -3 Then
                LeftRight = LeftRight - 0.2
            Else
                TurningLeft = False
            End If
            If Int(LeftRight) = 0 Then
                PlanePicX = 0
                PlaneChange = True
            End If
        Else
            If GetAsyncKeyState(vbKeyLeft) <> 0 Then
                If LeftRight >= -1 * MaxVel And LeftRight Then LeftRight = LeftRight - 0.1 'This adjust the planes velocity
                If DblClickLeft > 0 And DblClickLeft < 8 Then      'This checks to see if the user dbl clicked or not
                    TurningLeft = True                              'It just checks if the user has clicked twice within
                End If                                              '10 iterations of the while loop. It checks from 0 - 5
                DblClickLeft = 10                                   'because, a slow key press can be 5 iterations long
            ElseIf GetAsyncKeyState(vbKeyRight) <> 0 Then
                If LeftRight <= MaxVel Then LeftRight = LeftRight + 0.1   'Adjusts velocity
                If DblClickRight > 0 And DblClickRight < 8 Then
                    TurningRight = True
                End If
                DblClickRight = 10
            End If
            If GetAsyncKeyState(vbKeyUp) <> 0 Then
                If PlanePicY = 0 Then
                    If UpDown >= -7 Then
                        UpDown = UpDown - 0.2   'If the plane is facing up, it goes up quicker
                        MaxVel = 5
                    End If
                End If
                If UpDown >= -5 Then
                    UpDown = UpDown - 0.1
                End If
            ElseIf GetAsyncKeyState(vbKeyDown) <> 0 Then
                If PlanePicY = 160 Then
                    If UpDown <= 7 Then
                        UpDown = UpDown + 0.4
                    End If
                End If
                If UpDown <= 5 Then
                    UpDown = UpDown + 0.2
                End If
            ElseIf GetAsyncKeyState(vbKeyShift) <> 0 Then
                frmMain.Refresh
            Else
                If DblClickLeft > 0 Then DblClickLeft = DblClickLeft - 1
                If DblClickRight > 0 Then DblClickRight = DblClickRight - 1
            End If
        End If
        If LeftRight > MaxVel Then LeftRight = LeftRight - 0.1      ' This is in case the plane is going
        If LeftRight < -1 * MaxVel Then LeftRight = LeftRight + 0.1 ' to fast in a direction
        
        'This is a bit of aerodynamics for the plane
        If 4 - Abs(LeftRight) > 0 And TakeOff = False Then
            'Stall = (4 - LeftRight)       'This is the stall effect
            UpDown = UpDown + (4 - LeftRight) / 20
        Else
            If Stall > 0 Then Stall = Stall - 0.1
            If UpDown > 0 Then UpDown = UpDown - 0.01 ' Stall = -0.01    'This is a stabalizing effect
        End If
        
        BG_X = BG_X + LeftRight 'This part incrememnts the backgroun position
        TreeX = TreeX - LeftRight 'This part increments the tree position
        
        picPlane.Top = picPlane.Top + UpDown ' + Stall 'This moves the plane up and down
        If picPlane.Top > frmMain.ScaleHeight Then
            picPlane.Top = 200
            Stall = 0
            UpDown = 0
        End If
        'If PlaneChangeX = PlanePicX And PlaneChangeY = PlanePicY Then  ' The flicker reduction works well but I ran into the problem
        '    Blt (False)                                                ' that whenever it doesn't flicker, the plane moves MUCH
        'Else                                                           ' faster and throws everything off. If you think you can fix
            Blt (True)                                                  ' it, uncomment everything to the left of this text and the if
        'End If                                                         ' statement in the blt function (there's also a bit of comments
        'PlaneChangeX = PlanePicX                                       ' in the second if statement) That should get rid of the flicker
        'PlaneChangeY = PlanePicY                                       ' for the top half of the screen
    Loop Until GetAsyncKeyState(vbKeyEscape) <> 0
End Function

Private Sub Form_Terminate()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteGeneratedDC BG
    DeleteGeneratedDC PlaneMask
    DeleteGeneratedDC PlaneSprite
    DeleteGeneratedDC TreeMask
    DeleteGeneratedDC TreeSprite
    
    Set frmMain = Nothing
End Sub

Private Sub tmrBG_Timer()
    BG_X = BG_X + 5
    BitBlt frmMain.hdc, 0, 0, Me.Width, Me.Height, BG, BG_X, 0, vbSrcCopy
End Sub

Public Function Blt(PlaneChange As Boolean)
    'If PlaneChange = True Then '
    '    picPlane.Refresh       ' Put this if statement in for flicker reduction
    'End If                     '
    If picPlane.Top >= 500 - picPlane.Height Then 'And PlaneChange = False Then 'Put this part of the if statement in for flicker reduction
        BitBlt picPlane.hdc, 0, 0, picPlane.Width, picPlane.Height, BG, BG_X + picPlane.Left, picPlane.Top + 150, vbSrcCopy
        PlaneChange = True
    Else
        picPlane.Refresh  'Comment this out for flicker reduction
    End If
    If PlaneChange = True Then
        BitBlt picPlane.hdc, 0, 3, 120, 80, PlaneMask, PlanePicX, PlanePicY, vbSrcAnd
        BitBlt picPlane.hdc, 0, 3, 120, 80, PlaneSprite, PlanePicX, PlanePicY, vbSrcPaint
    End If
    'BitBlt frmMain.hdc, 50, 50, 120, 80, PlaneMask, PlanePicX, PlanePicY, vbSrcAnd
    'BitBlt frmMain.hdc, 50, 50, 120, 80, PlaneSprite, PlanePicX, PlanePicY, vbSrcPaint
    BitBlt frmMain.hdc, 0, 500, Me.ScaleWidth, Me.ScaleHeight - 500, BG, BG_X, 650, vbSrcCopy
'    If TreeX < Me.ScaleWidth Then
'        BitBlt frmMain.hdc, TreeX, 600, Me.ScaleWidth - TreeX, 80, TreeMask, 0, 0, vbSrcAnd
'        BitBlt frmMain.hdc, TreeX, 600, Me.ScaleWidth - TreeX, 80, TreeSprite, 0, 0, vbSrcPaint
'    End If
End Function
