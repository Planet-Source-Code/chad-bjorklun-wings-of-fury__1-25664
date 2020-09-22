Attribute VB_Name = "Vars"
Public BG As Long 'BackGround Picture

Public DrawPlaneMask As Long
Public DrawPlaneSprite As Long

Public PlaneMask As Long ' Picture of 6 planes with white background
Public PlaneSprite As Long ' Picture of 6 planes with black background

Public PlaneFireMask As Long
Public PlaneFireSprite As Long

Public TreeMask As Long
Public TreeSprite As Long

Public GuyMask As Long
Public GuySprite As Long

Public CarrierMask As Long
Public CarrierSprite As Long
Public CarrierMaskTemp As Long

Public GraveS_M As Long 'This is an example of another way to make sprites and masks. On this pic
                        'I put both on the same file. I could have done this on all of the files
                        'but since B&W files take up a lot less space, I though it would be best
                        'if they were all seperate. I only did it on this one since it is so small
                        
Public PlanePicX As Integer ' Draws the plane facing left or facing right
Public PlanePicY As Integer ' Draws the plane up, straight or down

'For you who don't understand BitBlt, look at this
'
'  0,0________________120,0________________     Here is an example of one of the plane bitmaps.
'    |                                     |    The horizontal and angled lines inside represent
'    |  \  \                         /  /  |    the plane and which way it is facing.  BitBlt lets
'    |    \  \                     /  /    |    you draw a rectangle around a section of a bmp by
'    |      \  \                 /  /      |    giving it the top left coords and specifying a
'0,80|               120,80                |    width and height.  So when the plane is flying up
'    |                                     |    and to the left, the coords are (0,0)  When the
'    |  -------                 -------    |    plane is flying straight right, the coords are
'    |  -------                 -------    |    (120,80)  The bmp is set up so that each rectangle
'    |                                     |    is the same width and height so that never has to
'0,160              120,160                |    change. In this case it will alway be 120 and 80
'    |      /  /                 \  \      |
'    |    /  /                     \  \    |
'    |  /  /                         \  \  |
'     _____________________________________
'
'One more thing for you BitBlt virgins, this coordinate system is done with two pictures, exactly
'the same except for color. When the two are combined, BitBlt checks for spots where the mask file
'has white and the sprite file has black. When the black and white overlap, bitblt makes it
'transparent. If you still don't understand, just look at the pictures.

Public BG_X As Integer  'As explained above, this is the x-coord of the bckgrnd file.  Only x is needed
                        'since the y never changes.  The background is a really long bmp, the program
                        'just draws enough to fill the screen

Public Plane_Y As Integer 'This is the y-coord where the plane is draw on the screen, since the
                          'background simulates the x movement, only a y-coord is needed

'The dounle click check could probably be done easier with a timer but I just really don't like timers
'so I did it my own way. And for all you timer lovers out there...you'll get what you deserve (I have
'no idea what that is supposed to mean)
Public DblClickLeft As Integer 'This variable will be used to check for a double click on the left key
Public DblClickRight As Integer 'Same as above but for right

Public LeftRight As Single  'These are the planes velocity in each direction
Public UpDown As Single     '

Public TurningLeft As Boolean  'These make sure that the plane completes its turn before the user can
Public TurningRight As Boolean 'use the keyboard

Public MaxVel As Single 'This is used because the maximum velocity changes depending on if the plane
                        'is facing up, down or level

Public PlaneChangeX As Integer  'These are used to help reduce flickering (sorry, I know it's still bad)
Public PlaneChangeY As Integer  'It makes it so the plane only redraws when it needs to, instead of
                                'every iteration of the loop
                                
Public TakeOff As Boolean   'This is used to make sure the program know it's okay to fly so slow

Public Stall As Single  'This variable produces a stall effect at slow speeds

Public TreeX As Integer 'This is the left side of the tree picture

Type People             '
    Visible As Boolean  '
    Alive As Boolean    'This is just the variable type I use for the people. Giving them
    x As Long           'there own parameters just makes it easier
    Speed As Integer    '
    PicPlace As Integer '
End Type                '

Public Guy(50) As People    'This represents the people that run around

Public CarrierX As Integer

Public UpKey As Integer     ' These are the keys used with GetAsyncKeyState. They are not necessary
Public DownKey As Integer   ' to use but it gives the user the option of defining his keys later on
Public LeftKey As Integer   ' instead of having them embedded in the code.  As a default, Up,Down,
Public RightKey As Integer  ' Left and Right are just the arrow keys (vbKeyUp...) The Gun key is
Public GunKey As Integer    ' shift(vbKeyShift) and the Bomb key is control(vbKeyControl) I didn't
Public BombKey As Integer   ' use Space or Enter because if the code breaks after using them, Visual
                            ' Studio will still implement them.
