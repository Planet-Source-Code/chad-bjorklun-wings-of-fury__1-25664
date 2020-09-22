VERSION 5.00
Begin VB.Form frmMainARTrue 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   Caption         =   "Form1"
   ClientHeight    =   11145
   ClientLeft      =   165
   ClientTop       =   420
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   743
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picGauges 
      Height          =   960
      Left            =   0
      Picture         =   "frmMainTrueAutoRedraw.frx":0000
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1020
      TabIndex        =   0
      Top             =   10305
      Width           =   15360
   End
   Begin VB.Label lblTemp3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Tow 
      Index           =   2
      X1              =   721
      X2              =   719
      Y1              =   638
      Y2              =   624
   End
   Begin VB.Line Tow 
      Index           =   1
      X1              =   755
      X2              =   753
      Y1              =   638
      Y2              =   624
   End
   Begin VB.Line Tow 
      Index           =   0
      X1              =   780
      X2              =   778
      Y1              =   638
      Y2              =   624
   End
   Begin VB.Shape shpCollision 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H000000FF&
      Height          =   495
      Left            =   3480
      Top             =   480
      Width           =   495
   End
   Begin VB.Label lblTemp2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblTemp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Image House 
      Height          =   375
      Index           =   4
      Left            =   3960
      Top             =   9240
      Width           =   255
   End
   Begin VB.Image House 
      Height          =   375
      Index           =   3
      Left            =   3600
      Top             =   9240
      Width           =   255
   End
   Begin VB.Image House 
      Height          =   375
      Index           =   2
      Left            =   3240
      Top             =   9240
      Width           =   255
   End
   Begin VB.Image House 
      Height          =   375
      Index           =   1
      Left            =   2880
      Top             =   9240
      Width           =   255
   End
   Begin VB.Image House 
      Height          =   375
      Index           =   0
      Left            =   2520
      Top             =   9240
      Width           =   255
   End
End
Attribute VB_Name = "frmMainARTrue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
'    MsgBox "ScaleHeight: " + Str(Me.ScaleHeight) + Chr(13) + "ScaleWidth: " + Str(Me.ScaleWidth)
'    MsgBox Me.ScaleTop
End Sub

Private Sub Form_DblClick()
    If Plane_Y <> 592 Then
        Call Form_Load
    Else
        Run_Game
    End If
End Sub


Private Sub Form_Load()
    Randomize
    'This if statement is here because sometimes I didn't want to load all the pictures because
    'I was just testing stuff
    If GetAsyncKeyState(vbKeyShift) = 0 Then
        'This is inside a while loop because sometimes the computer gets an
        'error while trying to load pictures. This just makes sure it keeps
        'trying until if finally gets it
        While BG <= 0
            BG = GenerateDC(App.Path + "\BGBigHSmall.bmp") 'BG now points at the spot in memory where the picture is placed
        Wend
    End If
    
    While PlaneMask <= 0
        PlaneMask = GenerateDC(App.Path + "\PlaneMask.bmp")
    Wend
    While PlaneSprite <= 0
        PlaneSprite = GenerateDC(App.Path + "\PlaneSprite.bmp")
    Wend
    
    While PlaneFireMask <= 0
        PlaneFireMask = GenerateDC(App.Path + "\PlaneFireMask.bmp")
    Wend
    While PlaneFireSprite <= 0
        PlaneFireSprite = GenerateDC(App.Path + "\PlaneFireSprite.bmp")
    Wend
    
    While TreeMask <= 0                                        'These are for trees, I was going to blt them on top of the
        TreeMask = GenerateDC(App.Path + "\TreesMask.bmp")     'bckgrnd so when/if the plane crashed because of the trees
    Wend                                                       'you would get the effect of the plain crashing behind the
                                                               'trees but it took up too much processor speed so I took it
    While TreeSprite <= 0                                      'out. If you have a fast enough processor, you might want to
        TreeSprite = GenerateDC(App.Path + "\TreesSprite.bmp") 'include this, just remember to calibrate it's positions
    Wend                                                       'using TreeX
    
    While GuyMask <= 0
        GuyMask = GenerateDC(App.Path + "\GuyMask.bmp")
    Wend
    While GuySprite <= 0
        GuySprite = GenerateDC(App.Path + "\GuySprite.bmp")
    Wend
    
    While GraveS_M <= 0
        GraveS_M = GenerateDC(App.Path + "\Grave.bmp")  'This file contains bothe the sprite and the mask
    Wend
    
    While CarrierMask <= 0
        CarrierMask = GenerateDC(App.Path + "\CarrierMask.bmp")
    Wend
    While CarrierSprite <= 0
        CarrierSprite = GenerateDC(App.Path + "\CarrierSprite.bmp")
    Wend
    While CarrierMaskTemp <= 0
        CarrierMaskTemp = GenerateDC(App.Path + "\CarrierMaskTemp.bmp")
    Wend
    
    DrawPlaneMasc = PlaneMask
    DrawPlaneSprite = PlaneSprite
    
    Plane_Y = 592 'Plane start being drawn near the ground
    
    PlanePicX = 120 '
    PlanePicY = 80  'This tells the program to draw the plane straight right
    
    BG_X = 0    'Backgroun starts at the beginning
    
    TakeOff = True  'I honestly don't know quite what I'm going to do with this yet
    
    MaxVel = 5  'This sets the maximum velocity of the plane
    
    UpDown = 0  'Makes the plane fly straight
    
    TreeX = 1925    'This is the x position of the Trees picture. All this is really here for
                    'is to create the effect of the plane flying behind the trees
    
    CarrierX = 100  'This is where the carrier is in relation to the start of the background
    
    House(0).Left = 3360    '
    House(0).Width = 23     '
                            '
    House(1).Left = 4695    ' These are just the positions of the houses.
    House(1).Width = 27     ' I used image boxes just because I was tired of so mane
                            ' variables. It makes it a little less accurate but you know
    House(2).Left = 6268    ' what, I don't give a damn.
    House(2).Width = 25     '
                            '
    House(3).Left = 7208    '
    House(3).Width = 26     '
                            '
    House(4).Left = 10060   '
    House(4).Width = 30     '
    
    'The following If statements set the position and speed of the people. I'm absolutly
    'posotive this could be done much easier with one for loop but as I said, this is a
    'work in progress in this came from old code. If you plan on making this into a
    'complete and finished project, I would definately fix this
    For g = 5 To 8
        Guy(g).Speed = Int(Rnd * 4) - 8
         If Guy(g).Speed >= 0 Then
            Guy(g).x = House(0).Left
        Else
            Guy(g).x = House(0).Left + House(0).Width
        End If
    Next

    For g = 9 To 15
        Guy(g).Speed = Int(Rnd * 8) - 4
        If Guy(g).Speed >= 0 Then
            Guy(g).x = House(1).Left
        Else
            Guy(g).x = House(1).Left + House(1).Width
        End If
    Next

    For g = 16 To 26
        Guy(g).Speed = Int(Rnd * 8) - 4
        If Guy(g).Speed >= 0 Then
            Guy(g).x = House(2).Left
        Else
            Guy(g).x = House(2).Left + House(2).Width
        End If
    Next

    For g = 27 To 38
        Guy(g).Speed = Int(Rnd * 8) - 4
        If Guy(g).Speed >= 0 Then
            Guy(g).x = House(3).Left
        Else
            Guy(g).x = House(3).Left + House(3).Width
        End If
    Next

    For g = 39 To 43
        Guy(g).Speed = Int(Rnd * 4) * -1
        If Guy(g).Speed >= 0 Then
            Guy(g).x = House(4).Left
        Else
            Guy(g).x = House(4).Left + House(4).Width
        End If
    Next

   ' For g = 44 To 50
   '     Guy(g).X = House(5).Left
   '     Guy(g).Speed = Int(Rnd * 4)
   ' Next

    For g = 5 To 43
        If Guy(g).Speed = 0 Then Guy(g).Speed = 1 'We want to make sure they move
        Guy(g).Alive = True 'They're obviously all alive to begin with
        Guy(g).Visible = True   'This changes whether they are in the house or up and runnin
    Next
    
    UpKey = vbKeyUp         'These are the default keys I'm using for testing. They are set
    DownKey = vbKeyDown     'to variables so that a user can later on define his own keys
    LeftKey = vbKeyLeft     'if he doesn't like the default
    RightKey = vbKeyRight   '
    GunKey = vbKeyShift     '
    BombKey = vbKeyControl  '
End Sub

Public Function Run_Game()
    'GetAsyncKeyState is a function that checks the state of a key, meaning it checks whether it is up or
    'down. If it is up(not pressed), the function returns 0, if it is down, it will return -32767 the first
    'time it checks the key. If the key is being held and the function keeps checking, it will incrememnt
    '-32767 by -1, it's rather confusing but just remember: If you want to check if it's simply being pushed
    'or not, use "<> 0" If you're checking for the first time it's pushed though, use "= -32767" Another way
    'to determine which to use is if the computer is checking it too fast, use the -32767, that should slow
    'it down and vice versa if it is too slow (If you're interested, -32767 is the lower bound of an integer)
    Do
        Me.Refresh  'Blanks the screen. Usually would make the animation flicker but AutoRedraw on the form is on
                    'This is necessary for BitBlt with or without AutoRedraw                             '

        If UpDown >= 0.65 Then      '
            PlanePicY = 160         ' This just checks to see how fast the plane is moving up and down, if
        ElseIf UpDown <= -1.5 And TakeOff = False Then  ' is is moving fast in either direction, it makes it so the plane is
            PlanePicY = 0           ' facing that direction. In other words, if the plane is going up steeply
        Else                        ' this will make it so it is facing upwards, instead of straight forward
            PlanePicY = 80          '
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
            If GetAsyncKeyState(LeftKey) <> 0 Then
                If LeftRight >= -1 * MaxVel And LeftRight Then LeftRight = LeftRight - 0.05 'This adjust the planes velocity
                If DblClickLeft > 7 And DblClickLeft < 10 Then      'This checks to see if the user dbl clicked or not
                    DblClickLeft = 0                                'It just checks if the user has clicked twice within
                    TurningLeft = True                              '10 iterations of the while loop. It checks from 0 - 5
                End If                                              'because, a slow key press can be 5 iterations long
                DblClickLeft = 10                                   '
            ElseIf GetAsyncKeyState(RightKey) <> 0 Then
                If LeftRight <= MaxVel Then LeftRight = LeftRight + 0.05   'Adjusts velocity
                If DblClickRight > 7 And DblClickRight < 10 Then
                    DblClickRight = 0
                    TurningRight = True
                End If
                DblClickRight = 10
            Else
                If DblClickLeft > 0 Then DblClickLeft = DblClickLeft - 1    'Acts as a timer for dbl click
                If DblClickRight > 0 Then DblClickRight = DblClickRight - 1 ' ^
            End If
            If GetAsyncKeyState(UpKey) <> 0 Then
                If PlanePicY = 0 Then
                    If UpDown >= -6 Then
                        UpDown = UpDown - Abs(LeftRight) / 15  'If the plane is facing up, it goes up quicker
                        MaxVel = 5
                    End If
                End If
                If UpDown >= -4 Then
                    UpDown = UpDown - Abs(LeftRight) / 28
                End If
            ElseIf GetAsyncKeyState(DownKey) <> 0 And TakeOff = False Then
                If PlanePicY = 160 Then
                    If UpDown <= 7 Then
                        UpDown = UpDown + 0.4
                    End If
                End If
                If UpDown <= 5 Then
                    UpDown = UpDown + 0.2
                End If
            Else
                If UpDown > 0 Then UpDown = UpDown - 0.01   'This is a stabalizing effect
                If UpDown < 0 Then UpDown = UpDown + 0.08
            End If
            If GetAsyncKeyState(GunKey) <> 0 Then
                If DrawPlaneMask = PlaneMask Then
                    DrawPlaneMask = PlaneFireMask
                    DrawPlaneSprite = PlaneFireSprite
                Else
                    DrawPlaneMask = PlaneMask
                    DrawPlaneSprite = PlaneSprite
                End If
            ElseIf GetAsyncKeyState(BombKey) <> 0 Then
                
            Else
                DrawPlaneMask = PlaneMask
                DrawPlaneSprite = PlaneSprite
            End If
        End If
        If LeftRight > MaxVel Then LeftRight = LeftRight - 0.1      ' This is in case the plane is going
        If LeftRight < -1 * MaxVel Then LeftRight = LeftRight + 0.1 ' to fast in a direction
        
        'This is a bit of aerodynamics for the plane
        If 4 - Abs(LeftRight) > 0 And TakeOff = False Then
            If TurningLeft = True Or TurningRight = True Then
                UpDown = UpDown + (4.5 - LeftRight) / 50
            Else
                'UpDown = UpDown + (6 - LeftRight) / 35  'This is the stall effect
                 UpDown = UpDown + Abs(LeftRight) / 28 + (5 - Abs(LeftRight)) / 10
            End If
        End If
        
        'This is * 4 because I had originally done this without the use of AutoRedraw
        'but since I've decided to use it I have to speed things up because AutoRedraw
        'makes everything slower. It does make it a bit jumpier but I think it's benifits
        'outweigh the drawbacks
        BG_X = BG_X + LeftRight * 4 'This part incrememnts the backgroun position
        'The other pictures have to move with the background too
        TreeX = TreeX - LeftRight * 4 'This part increments the tree position
        CarrierX = CarrierX - LeftRight * 4 'Carrier (If you couldn't guess)
        For h = 0 To 4
            House(h).Left = House(h).Left - LeftRight * 4   'Houses
        Next h
        For t = 0 To 2                              'This is used to move the Tow lines (the lines the catch the
            Tow(t).x1 = Tow(t).x1 - LeftRight * 4   'plane on the carrier)
            Tow(t).x2 = Tow(t).x2 - LeftRight * 4   '
        Next t                                      '

        Plane_Y = Plane_Y + UpDown * 2 'This moves the plane up and down on the screen
        
        If Plane_Y > Me.ScaleHeight Then   ' I used this while I was making the game,
            Plane_Y = 200                       ' I was tired of losing the plane below the
            Stall = 0                           ' screen and having to reload the game
            UpDown = 0                          '
        End If                                  '
        
        'This for statement is for all the people running around
        For g = 0 To 50
            Guy(g).x = Guy(g).x - LeftRight * 4 'This moves the people with the background
            If Guy(g).Alive = True Then
                Guy(g).x = Guy(g).x - Guy(g).Speed * 4 'This makes them walk
                Guy(g).PicPlace = Guy(g).PicPlace + 20              'This is for their animation
                If Guy(g).PicPlace = 40 Then Guy(g).PicPlace = 0    '
                For h = 0 To 4
                    'This checks if they have ran into a house
                    If Guy(g).x + 2 > House(h).Left And Guy(g).x < House(h).Left + House(h).Width Then
                        If h = 0 Then                                   'This makes sure if they are at
                            Guy(g).Speed = Int(Rnd * 4) * -1            'the first house, they have to
                            If Guy(g).Speed = 0 Then Guy(g).Speed = -1  'run to the right
                        ElseIf h = 4 Then                               '
                            Guy(g).Speed = Int(Rnd * 4)                 'This makes them run left
                            If Guy(g).Speed = 0 Then Guy(g).Speed = 1   '
                        Else
                            Guy(g).Speed = Int(Rnd * 8) - 4             'This is for all the in between houses
                        End If
                        If Guy(g).Speed >= 1 Then                       'This just puts them on the left of the
                            Guy(g).x = House(h).Left                    'house if they're running left and on
                        Else                                            'the right if they're running to the right
                            Guy(g).x = House(h).Left + House(h).Width   '
                        End If                                          '
                        'Guy(g).Visible = False
                    End If
                Next h
            End If
        Next g
        Blt (True)  'Calls the function the draws everything on the screen
lblTemp.Caption = DblClickLeft
lblTemp2.Caption = DblClickRight
lblTemp3.Caption = Plane_Y
        'Loops until you hit Escape
        Loop Until GetAsyncKeyState(vbKeyEscape) <> 0
End Function

Private Sub Form_Terminate()
    'This part deletes all the pictures that were stored in memory. This is VERY inportant
    'to do if you ever use Memory DC's. If you don't, after a while your computer will
    'lock up on you because it is out of memory. It doesn't permanently hurt your
    'computer or anything, you just have to restart to get the memory back. When this
    'happens, it's called Memory Leakage.
    DeleteGeneratedDC BG
    DeleteGeneratedDC PlaneMask
    DeleteGeneratedDC PlaneSprite
    DeleteGeneratedDC TreeMask
    DeleteGeneratedDC TreeSprite
    DeleteGeneratedDC GuyMask
    DeleteGeneratedDC GuySprite
    DeleteGeneratedDC GraveS_M
    DeleteGeneratedDC CarrierMask
    DeleteGeneratedDC CarrierSprite
    
    Unload Me
    Set frmMainARTrue = Nothing
End Sub


Private Sub picGauges_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = 1 And Shift = 0 Then picGauges.Width = picGauges.Width + 1
'    If Button = 2 And Shift = 0 Then picGauges.Width = picGauges.Width - 1
'    If Button = 1 And Shift = 1 Then picGauges.Height = picGauges.Height + 1
'    If Button = 2 And Shift = 1 Then picGauges.Height = picGauges.Height - 1
'    MsgBox "Width: " + Str(picGauges.Width) + Chr(13) + "Height: " + Str(picGauges.Height) + Chr(13) + "Top: " + Str(picGauges.Top)
End Sub



Public Function Blt(PlaneChange As Boolean)
'BITBLT
'---------
'BitBlt Destination, X-Coord on Dest, Y-Coord on Dest, Width to be Drawn, Height to be Drawn, Picture to Draw, X-Coord on Pic to Draw from, Y-Coord, How to Draw
'Destination:   This is where you want to draw your picture. Usually the only two you will uses is either the form of a picture box (Not an image box)
'                Just make sure you include the '.hdc'
'X-Coord and Y-Coord on Dest:   This is where you want to draw the pic on the Destination. 0,0 is the top left.
'Width and Height to be Drawn:  Pretty much self explanatory
'Picture to be Drawn From:      This is the picture you want to draw. It can either be a memory DC, like I've used or it can be an object, like a form a picture
'                               box that has a 'hdc' property.
'X-Coord and Y-Coord on Pic:    This is where you want to start drawing from on the picture. It will use the Width and Height from before. 0,0 is top left
'How to Draw:   This will make the function do different things. There are a lot of things it can do but I'lll just show you a few
'               vbSrcCopy:  This just draws the picture, nothing fancy
'               vbSrcAnd:   Draws picture on screen using an AND boolean operation
'               vbSrcPaint: Draws picture using an OR bollean operation
'               When vbSrcPaint is used after vbSrcAnd, and the right pictures are used, you can make certain parts of the picture transparent. See Vars module for more info
'----------

'This draws the BackGround
    BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight - 62, BG, BG_X, 150, vbSrcCopy
'This draws the Carrier
    If CarrierX + 884 >= 0 And CarrierX <= Me.ScaleWidth Then
        BitBlt Me.hdc, CarrierX, 530, 884, 150, CarrierMask, 0, 0, vbSrcAnd
        BitBlt Me.hdc, CarrierX, 530, 884, 150, CarrierSprite, 0, 0, vbSrcPaint
    End If
'This draws the plane
    BitBlt Me.hdc, Me.ScaleWidth / 2 - 60, Plane_Y, 120, 80, DrawPlaneMask, PlanePicX, PlanePicY, vbSrcAnd
    BitBlt Me.hdc, Me.ScaleWidth / 2 - 60, Plane_Y, 120, 80, DrawPlaneSprite, PlanePicX, PlanePicY, vbSrcPaint
'This draws the people and the graves
    For g = 0 To 50
        If Guy(g).x >= 0 And Guy(g).x <= Me.ScaleWidth And Guy(g).Visible = True Then
            If Guy(g).Alive = True Then
                BitBlt Me.hdc, Guy(g).x, 640, 20, 30, GuyMask, Guy(g).PicPlace, 0, vbSrcAnd
                BitBlt Me.hdc, Guy(g).x, 640, 20, 30, GuySprite, Guy(g).PicPlace, 0, vbSrcPaint
            Else
                BitBlt Me.hdc, Guy(g).x, 650, 20, 30, GraveS_M, 15, 0, vbSrcAnd
                BitBlt Me.hdc, Guy(g).x, 650, 20, 30, GraveS_M, 0, 0, vbSrcPaint
            End If
        End If
    Next g
'This draws the Trees
    If TreeX >= -1300 And TreeX <= Me.ScaleWidth Then
        BitBlt Me.hdc, TreeX, 615, 1300, 80, TreeMask, 0, 0, vbSrcAnd
        BitBlt Me.hdc, TreeX, 615, 1300, 80, TreeSprite, 0, 0, vbSrcPaint
    End If
    

    If CollisionDetect(Me.ScaleWidth / 2 - 60, Plane_Y + 4, 120, 80, PlanePicX, PlanePicY, PlaneMask, CarrierX, 642, 864, 43, 0, 0, CarrierMaskTemp) Then
        shpCollision.FillStyle = 0
        If TakeOff = False Then
            If Plane_Y > 597 And CarrierX + 838 < Me.ScaleWidth / 2 - 60 Then
                LeftRight = 0
            
            ElseIf Abs(UpDown) > 0.5 Then
                UpDown = 0
            Else
                TakeOff = True
            End If
        Else
            
        End If
    Else
        TakeOff = False
        shpCollision.FillStyle = 1
    End If
    
'Order is important when using Bitblt, it creates the perception of depth.
End Function
