VERSION 5.00
Begin VB.Form Mainfrm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   180
   ClientTop       =   0
   ClientWidth     =   12000
   ForeColor       =   &H0000FF00&
   LinkTopic       =   "Form1"
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Mainfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'article Purpose:
'The point of this code is to show anyone interested then
'Power of bit blt. This is just a short code, and I tried to comment well.
'This is not supposed to be advanced. Just advanced enough
'to be amusing and hopefully helpful. Feel free to copy/cut/paste
'Just give credit where it is due, and enjoy
'[MAINFRM]Purpose:
'To allow user to control his/her ship and to allow user to
'see his/her ship and the enemy's


Option Explicit  'forces all variables to be declared
'the bitblt image stuff api
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'system timer
Private Declare Function GetTickCount Lib "kernel32" () As Long
'turns things into bitmaps for bitblt
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'makes device contexts for bitblt
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Gets current device context
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'loads/selects device context/object
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'removes device context/object from memory/
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'hides mouse
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'tests keyboard keyboard for keys up/down
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
'my udt
'declaring  your own type helps organize your code
'If you want to use a lot of udts you should use
'classes instead, udts sometimes are weird with memory
Private Type Space_Ship
    X As Single
    Y As Single
    Width As Integer
    Height As Integer
    XSpeed As Integer
    YSpeed As Integer
    Direction As Byte
    LPic As Long
    RPic As Long
    UPic As Long
    DPic As Long
    ShieldPic As Long
End Type

Private BufferBMP0 As Long 'the buffer for bitblt
Private BackBuffer0 As Long 'more bitblt stuff
Private Player As Space_Ship, Enemy As Space_Ship 'player and enemy
Private a As Byte 'loop key

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Load/Unload----------------------------------------
'///////////////////////////////////////////////////
Private Sub Form_Load()
    Dim T1 As Long, T2 As Long 'Tick count variables
    SetUpWorld  'setup procedure
    T2 = GetTickCount 'get the value from the system clock
    Show  'Force the form to be drawn
    ShowCursor 0
    Do While a = 1   '"a" is the unload key
        DoEvents  'Allow other procedures to happen
        Show    'Force the form to be redrawn
        T1 = GetTickCount
        If (T1 - T2) >= 15 Then 'Wait 15 ms before anything happens
            DoMainLoop    'call to the main procedure
            T2 = GetTickCount   'refresh the value from the system clock
        End If
    Loop
    ShowCursor 1
    Unload Me 'when the loop stops immediately unload
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'remove all objects from memory
    DeleteObject BufferBMP0
    DeleteDC BackBuffer0
    With Player
        DeleteDC .UPic
        DeleteDC .DPic
        DeleteDC .LPic
        DeleteDC .RPic
        DeleteDC .ShieldPic
    End With
    With Enemy
        DeleteDC .UPic
        DeleteDC .DPic
        DeleteDC .LPic
        DeleteDC .RPic
        DeleteDC .ShieldPic
    End With
    End  'close app
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Key Events-----------------------------------------
'///////////////////////////////////////////////////
Private Sub CheckKeys()
    'Keypressed() and the api that goes with it is great
    'With it you can test for multiple keys being pressed
    'at one time, plus it is simple to use
    If KeyPressed(vbKeyEscape) Then a = 254
    With Player 'the following code is for accelerating the player
        If KeyPressed(vbKeyUp) Then
            '.Direction = 0 'i use this to tell which picture to show
            .YSpeed = .YSpeed - 1
            If .YSpeed < -10 Then .YSpeed = -20 'speed limit
        End If
        If KeyPressed(vbKeyDown) Then
           ' .Direction = 1
            .YSpeed = .YSpeed + 1
            If .YSpeed > 10 Then .YSpeed = 20
        End If
        If KeyPressed(vbKeyLeft) Then
           ' .Direction = 2
            .XSpeed = .XSpeed - 1
            If .XSpeed < -10 Then .XSpeed = -20
        End If
        If KeyPressed(vbKeyRight) Then
           ' .Direction = 3
            .XSpeed = .XSpeed + 1
            If .XSpeed > 10 Then .XSpeed = 20
        End If
        'this is for shooting the gun [space][enter]or[control]
        If KeyPressed(vbKeyReturn) Or KeyPressed(vbKeySpace) Or KeyPressed(vbKeyControl) Then
            If .Direction = 0 Then
                Line (.X + .Width / 4, .Y)-(.X + .Width / 4, 0), vbGreen
                Line (.X + .Width * 3 / 4, .Y)-(.X + .Width * 3 / 4, 0), vbGreen
                If .Y > Enemy.Y And ((.X + .Width / 4 >= Enemy.X And .X + .Width / 4 <= Enemy.X + Enemy.Width) Or (.X + .Width * 3 / 4 >= Enemy.X And .X + .Width * 3 / 4 <= Enemy.X + Enemy.Width)) Then BitBlt BackBuffer0, Enemy.X, Enemy.Y, Enemy.Width, Enemy.Height, Enemy.ShieldPic, 0, 0, vbSrcPaint
            ElseIf .Direction = 1 Then
                Line (.X + .Width / 4, .Y + .Height)-(.X + .Width / 4, Me.ScaleHeight), vbGreen
                Line (.X + .Width * 3 / 4, .Y + .Height)-(.X + .Width * 3 / 4, Me.ScaleHeight), vbGreen
                If .Y < Enemy.Y And ((.X + .Width / 4 >= Enemy.X And .X + .Width / 4 <= Enemy.X + Enemy.Width) Or (.X + .Width * 3 / 4 >= Enemy.X And .X + .Width * 3 / 4 <= Enemy.X + Enemy.Width)) Then BitBlt BackBuffer0, Enemy.X, Enemy.Y, Enemy.Width, Enemy.Height, Enemy.ShieldPic, 0, 0, vbSrcPaint
            ElseIf .Direction = 2 Then
                Line (.X, .Y + .Height / 4)-(0, .Y + .Height / 4), vbGreen
                Line (.X, .Y + .Height * 3 / 4)-(0, .Y + .Height * 3 / 4), vbGreen
                If .X > Enemy.X And ((.Y + .Height / 4 >= Enemy.Y And .Y + .Height / 4 <= Enemy.Y + Enemy.Height) Or (.Y + .Height * 3 / 4 >= Enemy.Y And .Y + .Height * 3 / 4 <= Enemy.Y + Enemy.Height)) Then BitBlt BackBuffer0, Enemy.X, Enemy.Y, Enemy.Width, Enemy.Height, Enemy.ShieldPic, 0, 0, vbSrcPaint
            ElseIf .Direction = 3 Then
                Line (.X + .Width, .Y + .Height / 4)-(Me.ScaleWidth, .Y + .Height / 4), vbGreen
                Line (.X + .Width, .Y + .Height * 3 / 4)-(Me.ScaleWidth, .Y + .Height * 3 / 4), vbGreen
                If .X < Enemy.X And ((.Y + .Height / 4 >= Enemy.Y And .Y + .Height / 4 <= Enemy.Y + Enemy.Height) Or (.Y + .Height * 3 / 4 >= Enemy.Y And .Y + .Height * 3 / 4 <= Enemy.Y + Enemy.Height)) Then BitBlt BackBuffer0, Enemy.X, Enemy.Y, Enemy.Width, Enemy.Height, Enemy.ShieldPic, 0, 0, vbSrcPaint
            End If
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'KeyCode = last key pressed
    With Player 'i use this to tell which direction to face
        If KeyCode = (vbKeyUp) Then .Direction = 0 'i use this to tell which picture to show
        If KeyCode = (vbKeyDown) Then .Direction = 1
        If KeyCode = (vbKeyLeft) Then .Direction = 2
        If KeyCode = (vbKeyRight) Then .Direction = 3
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'MAIN Loop------------------------------------------
'///////////////////////////////////////////////////
Private Sub DoMainLoop()
    ClearImages 'procedure-erase images
    CalculatePosition  'do any math or calculations
    MakeEnemyDoStuff 'make the enemy do his thing
    ShowPlayers 'paint images onto form
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Pathing--------------------------------------------
'///////////////////////////////////////////////////
Private Sub CalculatePosition()
    With Player
        Dim a As Integer, b As Integer 'temporary values that i will put the x/yspeeds into so i can calculate stuff easier
        'the following four lines loop the player back onto the screen
        'if it strays off the edge
        If .X > Me.ScaleWidth Then .X = 0
        If .X + .Width < 0 Then .X = Me.ScaleWidth - .Width
        If .Y > Me.ScaleHeight Then .Y = 0
        If .Y + .Height < 0 Then .Y = Me.ScaleHeight - .Height
        'change position
        .Y = .Y + .YSpeed / 2
        .X = .X + .XSpeed / 2
        'Collision Detection------
        If Dist(.X + .Width / 2, .Y + .Height / 2, Enemy.X + Enemy.Width / 2, Enemy.Y + Enemy.Height / 2) <= .Width Then
            .Y = .Y + .YSpeed * -2
            .X = .X + .XSpeed * -2
            .YSpeed = -.YSpeed
            .XSpeed = -.XSpeed 'make the ships bounce off each other
            Enemy.Y = Enemy.Y + Enemy.YSpeed * -2
            Enemy.X = Enemy.X + Enemy.XSpeed * -2
            a = (Enemy.YSpeed - .YSpeed) / 2
            b = (Enemy.XSpeed - .XSpeed) / 2
            .YSpeed = (.YSpeed + Enemy.YSpeed) / 2
            .XSpeed = (.XSpeed + Enemy.XSpeed) / 2
            Enemy.YSpeed = a
            Enemy.XSpeed = b
            BitBlt BackBuffer0, Enemy.X, Enemy.Y, Enemy.Width, Enemy.Height, Enemy.ShieldPic, 0, 0, vbSrcPaint
            BitBlt BackBuffer0, .X, .Y, .Width, .Height, .ShieldPic, 0, 0, vbSrcPaint
        End If
        '-----
    End With
    ClearImages
    With Enemy
        'the following four lines loop the enemy back onto the screen
        'if it strays off the edge
        If .X > Me.ScaleWidth Then .X = 0
        If .X + .Width < 0 Then .X = Me.ScaleWidth - .Width
        If .Y > Me.ScaleHeight Then .Y = 0
        If .Y + .Height < 0 Then .Y = Me.ScaleHeight - .Height
        'change position
        .Y = .Y + .YSpeed / 2
        .X = .X + .XSpeed / 2
    End With
    CheckKeys  'this sub is where I accept keyboard input
End Sub

Private Sub MakeEnemyDoStuff()
    Dim rndmov As Integer
    rndmov = Int(Rnd * 25)
    With Enemy
        'this conditional makes the enemy seem to fly his ship
        If rndmov = 7 Then
            .Direction = Int(Rnd * 4)
            .XSpeed = Int(Rnd * 41) - 20
            .YSpeed = Int(Rnd * 41) - 20
        End If
        'the following conditionals make the enemy shoot
        'at the player
        If .Y > Player.Y And ((.X + .Width / 4 >= Player.X And .X + .Width / 4 <= Player.X + Player.Width) Or (.X + .Width * 3 / 4 >= Player.X And .X + .Width * 3 / 4 <= Player.X + Player.Width)) Then
            .Direction = 0
            BitBlt BackBuffer0, Player.X, Player.Y, Player.Width, Player.Height, Player.ShieldPic, 0, 0, vbSrcPaint
            Line (.X + .Width / 4, .Y)-(.X + .Width / 4, 0), RGB(255, 0, 255)
            Line (.X + .Width * 3 / 4, .Y)-(.X + .Width * 3 / 4, 0), RGB(255, 0, 255)
        End If
        If .Y < Player.Y And ((.X + .Width / 4 >= Player.X And .X + .Width / 4 <= Player.X + Player.Width) Or (.X + .Width * 3 / 4 >= Player.X And .X + .Width * 3 / 4 <= Player.X + Player.Width)) Then
            .Direction = 1
            Line (.X + .Width / 4, .Y + .Height)-(.X + .Width / 4, Me.ScaleHeight), RGB(255, 0, 255)
            Line (.X + .Width * 3 / 4, .Y + .Height)-(.X + .Width * 3 / 4, Me.ScaleHeight), RGB(255, 0, 255)
            BitBlt BackBuffer0, Player.X, Player.Y, Player.Width, Player.Height, Player.ShieldPic, 0, 0, vbSrcPaint
        End If
        If .X > Player.X And ((.Y + .Height / 4 >= Player.Y And .Y + .Height / 4 <= Player.Y + Player.Height) Or (.Y + .Height * 3 / 4 >= Player.Y And .Y + .Height * 3 / 4 <= Player.Y + Player.Height)) Then
            .Direction = 2
            Line (.X, .Y + .Height / 4)-(0, .Y + .Height / 4), RGB(255, 0, 255)
            Line (.X, .Y + .Height * 3 / 4)-(0, .Y + .Height * 3 / 4), RGB(255, 0, 255)
            BitBlt BackBuffer0, Player.X, Player.Y, Player.Width, Player.Height, Player.ShieldPic, 0, 0, vbSrcPaint
        End If
        If .X < Player.X And ((.Y + .Height / 4 >= Player.Y And .Y + .Height / 4 <= Player.Y + Player.Height) Or (.Y + .Height * 3 / 4 >= Player.Y And .Y + .Height * 3 / 4 <= Player.Y + Player.Height)) Then
            .Direction = 3
            Line (.X + .Width, .Y + .Height / 4)-(Me.ScaleWidth, .Y + .Height / 4), RGB(255, 0, 255)
            Line (.X + .Width, .Y + .Height * 3 / 4)-(Me.ScaleWidth, .Y + .Height * 3 / 4), RGB(255, 0, 255)
            BitBlt BackBuffer0, Player.X, Player.Y, Player.Width, Player.Height, Player.ShieldPic, 0, 0, vbSrcPaint
        End If
    End With
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Show Players---------------------------------------
'///////////////////////////////////////////////////
Private Sub ClearImages()
    With Player
        If .Direction = 0 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 1 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 2 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 3 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
    End With
    With Enemy 'erase images
        If .Direction = 0 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 1 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 2 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
        If .Direction = 3 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, 0, 0, 0, vbBlackness
    End With
End Sub

Private Sub ShowPlayers()
    With Player 'show images
        If .Direction = 0 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .UPic, 0, 0, vbSrcPaint
        If .Direction = 1 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .DPic, 0, 0, vbSrcPaint
        If .Direction = 2 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .LPic, 0, 0, vbSrcPaint
        If .Direction = 3 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .RPic, 0, 0, vbSrcPaint
        If KeyPressed(vbKeyTab) Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .ShieldPic, 0, 0, vbSrcPaint
    End With
    With Enemy
        If .Direction = 0 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .UPic, 0, 0, vbSrcPaint
        If .Direction = 1 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .DPic, 0, 0, vbSrcPaint
        If .Direction = 2 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .LPic, 0, 0, vbSrcPaint
        If .Direction = 3 Then BitBlt BackBuffer0, .X, .Y, .Width, .Height, .RPic, 0, 0, vbSrcPaint
    End With
    'resetup the back buffer
    BitBlt Me.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, BackBuffer0, 0, 0, vbSrcCopy
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Initialize SS--------------------------------------
'///////////////////////////////////////////////////
Private Sub SetUpWorld()
    Randomize 'setup the random number generator
    'make scalesize=screensize
    Me.ScaleWidth = Screen.Width / Screen.TwipsPerPixelX
    Me.ScaleHeight = Screen.Height / Screen.TwipsPerPixelY
    'bitblt stuff to set up the buffers
    BackBuffer0 = CreateCompatibleDC(GetDC(0))
    BufferBMP0 = CreateCompatibleBitmap(GetDC(0), Me.ScaleWidth, Me.ScaleHeight)
    SelectObject BackBuffer0, BufferBMP0
    BitBlt BackBuffer0, 0, 0, Me.ScaleWidth, Me.ScaleHeight, 0, 0, 0, vbWhiteness
    With Player
        'load player's graphics into memory
        .UPic = LoadGraphicDC(App.Path & "\Images\Up.bmp")
        .DPic = LoadGraphicDC(App.Path & "\Images\Down.bmp")
        .LPic = LoadGraphicDC(App.Path & "\Images\Left.bmp")
        .RPic = LoadGraphicDC(App.Path & "\Images\Right.bmp")
        .ShieldPic = LoadGraphicDC(App.Path & "\Images\Shield.bmp")
        .X = 380
        .Y = 280
        .Width = 40
        .Height = 40
        .Direction = 0
        .XSpeed = 0
        .YSpeed = 0
    End With
    With Enemy
        'load Enemies's graphics into memory
        .UPic = LoadGraphicDC(App.Path & "\Images\Eup.bmp")
        .DPic = LoadGraphicDC(App.Path & "\Images\Edown.bmp")
        .LPic = LoadGraphicDC(App.Path & "\Images\Eleft.bmp")
        .RPic = LoadGraphicDC(App.Path & "\Images\Eright.bmp")
        .ShieldPic = LoadGraphicDC(App.Path & "\Images\Shield.bmp")
        .X = Int(Rnd * Me.ScaleWidth + 1)
        .Y = Int(Rnd * Me.ScaleHeight + 1)
        .Width = 40
        .Height = 40
        .Direction = 0
        .XSpeed = Int(Rnd * 41) - 20
        .YSpeed = Int(Rnd * 41) - 20
    End With
    a = 1 'this is the key to run the SS
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Functions------------------------------------------
'///////////////////////////////////////////////////
Private Function LoadGraphicDC(FileName0 As String) As Long
    'this function loads the graphic into memory
    On Error Resume Next 'simple error handling
    Dim TLoadGraphicDCT As Long
    TLoadGraphicDCT = CreateCompatibleDC(GetDC(0))
    SelectObject TLoadGraphicDCT, LoadPicture(FileName0)
    LoadGraphicDC = TLoadGraphicDCT
End Function

Private Function KeyPressed(KeyCode As Long) As Boolean
    'This function works with the api for keyboard input
    If GetKeyState(KeyCode) < -125 Then
        KeyPressed = True
    End If
End Function

Private Function Dist(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Double
    'the distance formula
    Dist = Sqr(((X2 - X1) * (X2 - X1)) + ((Y2 - Y1) * (Y2 - Y1)))
End Function

Rem ember: Made by Brian Adriance
