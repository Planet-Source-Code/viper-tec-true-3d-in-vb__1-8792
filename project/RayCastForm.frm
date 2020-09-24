VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form RayCastForm 
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   5940
   Icon            =   "RayCastForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   396
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1080
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer FPSTimer 
      Interval        =   1000
      Left            =   4590
      Top             =   3240
   End
   Begin VB.Timer RenderTimer 
      Interval        =   1
      Left            =   405
      Top             =   3240
   End
   Begin VB.Image OverlayImage 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   4080
      Top             =   3480
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image UnderlayImage 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   3240
      Top             =   3480
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image TexturesImage 
      Appearance      =   0  'Flat
      Height          =   555
      Left            =   2280
      Top             =   3510
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image ViewportImage 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5376
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Use the arrow keys to move - H = Home, Shift = Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   3240
      Width           =   4740
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu LoadMap 
         Caption         =   "Load Map"
      End
      Begin VB.Menu EditMap 
         Caption         =   "Edit Map"
      End
      Begin VB.Menu SaveMap 
         Caption         =   "Save Map"
      End
      Begin VB.Menu FileNull2 
         Caption         =   "-"
      End
      Begin VB.Menu LoadTextures 
         Caption         =   "Load Textures"
      End
      Begin VB.Menu FileNull1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Color 
      Caption         =   "Color"
      Begin VB.Menu UseColor 
         Caption         =   "Use Color"
         Checked         =   -1  'True
      End
      Begin VB.Menu UseTexture 
         Caption         =   "Use Textures"
      End
      Begin VB.Menu ColorNull1 
         Caption         =   "-"
      End
      Begin VB.Menu SetColors 
         Caption         =   "Set Colors"
      End
   End
   Begin VB.Menu Light 
      Caption         =   "Light"
      Begin VB.Menu LightOn 
         Caption         =   "Light On"
         Checked         =   -1  'True
      End
      Begin VB.Menu LightOff 
         Caption         =   "Light Off"
      End
      Begin VB.Menu LightNull1 
         Caption         =   "-"
      End
      Begin VB.Menu GammaLevel 
         Caption         =   "Gamma Level 1"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu GammaLevel 
         Caption         =   "Gamma Level 2"
         Index           =   2
      End
      Begin VB.Menu GammaLevel 
         Caption         =   "Gamma Level 3"
         Index           =   3
      End
      Begin VB.Menu GammaLevel 
         Caption         =   "Gamma Level 4"
         Index           =   4
      End
      Begin VB.Menu GammaLevel 
         Caption         =   "Gamma Level 5"
         Index           =   5
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "RayCastForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
Dim OldPlayerX As Integer       ' Old player's X position (for bounds check)
Dim OldPlayerY As Integer       ' Old player's Y position (for bounds check)
'
Dim ALL_STOP As Boolean         ' Rendering loop control flag
Dim FPS As Integer              ' Frames-Per-Second counter
Private Sub About_Click()
  '
  Call ViewportImage_Click
  '
End Sub

Private Sub EditMap_Click()
  '
  EditMapForm.Show 1
  '
End Sub

Private Sub Exit_Click()
  '
  Unload Me
  '
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  '
  ' Fill key array with keycode (changes suggested by Gary Allen Beebe - 8-24-99)
  '
  If KeyCode < 101 Then PlyrKeys(KeyCode) = True
  '
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  '
  ' Empty key array of keycode (changes suggested by Gary Allen Beebe - 8-24-99)
  '
  If KeyCode < 101 Then Let PlyrKeys(KeyCode) = False
  '
End Sub
Private Sub Form_Load()
  '
  ' Set up lighting
  '
  USE_COLOR = True: FORE_COLOR = 255: FADE_TO_COL = 0
  FLOOR_COLOR = 112: CEILING_COLOR = 99
  '
  Call GammaLevel_Click(5)
  Call LightOff_Click
  '
  ' Build translation table for lighting effect
  '
  Call BuildFogTransTable
  '
  ' Build table of Arctangents
  '
  Call BuildAtnTable
  '
  ' Initialize the textures
  '
  Call PictArrayInit1D(TexturesImage, App.Path + "\textures.bmp", sa2, bmp2, textures())
  Call PictArrayInit1D(UnderlayImage, App.Path + "\underlay.bmp", sa3, bmp3, underlay())
  Call PictArrayInit1D(OverlayImage, App.Path + "\overlay.bmp", sa4, bmp4, overlay())
  '
  ' Load the map for the maze
  '
  Call Load_Map(App.Path + "\default.maz")
  '
  ' Set up initial player position
  '
  ViewAngle = PlayerHomeA
  PlayerX = PlayerHomeX
  PlayerY = PlayerHomeY
  '
  ' Reset render loop flag and frame counter
  '
  ALL_STOP = False: FPS = 0
  '
End Sub
Private Sub Form_Resize()
  '
  ' I decided to add a resizing option after seeing an
  ' executable only version of this demo submitted to me
  ' by a person going by TomB or TR2 (his web page is
  ' "http://home.wish.net/~tomb/VB"). His was a full
  ' screen version - I wanted to make mine a little more
  ' flexible - so this is what I came up with...
  '
  ' Make sure form is visible (and not minimized)
  '
  If RayCastForm.ScaleWidth > 0 And RayCastForm.ScaleHeight > 0 Then
    '
    ' Only allow resize of image if ScaleWidth is big enough
    '
    If RayCastForm.ScaleWidth > 0 Then
      '
      ViewportImage.Width = RayCastForm.ScaleWidth
      '
      ' Position help label
      '
      Label1.Left = (ViewportImage.Width \ 2) - (Label1.Width \ 2)
      '
    End If
    '
    ' Only allow resize if ScaleHeight is tall enough
    '
    If RayCastForm.ScaleHeight - 50 > 0 Then
      '
      ViewportImage.height = RayCastForm.ScaleHeight - 50
      '
      ' Position help label
      '
      Label1.top = ViewportImage.height + 20
      '
    End If
    '
    ' (Re)Initialize the viewport (array pointer to bitmap)
    '
    ' When form is loaded, the form's Resize event is called,
    ' properly initializing the image buffer first thing...
    '
    Call PictArrayInit1D(ViewportImage, App.Path + "\blank.bmp", sa1, bmp1, viewport())
    '
  End If
  '
End Sub
Private Sub Form_Unload(Cancel As Integer)
  '
  ' Set flag to halt rendering loop
  '
  ALL_STOP = True
  '
  ' Destroy pointer to bitmap array and free up memory
  '
  Call PictArrayKill(viewport())
  Call PictArrayKill(textures())
  '
End Sub
Private Sub GammaLevel_Click(Index As Integer)
  '
  For t% = 1 To 5
    '
    If t% = Index Then
      GammaLevel(t%).Checked = True
      GAMMA = t%
    Else
      GammaLevel(t%).Checked = False
    End If
    '
  Next
  '
End Sub
Private Sub LightOff_Click()
  '
  FOG_ENABLE = False
  '
  LightOn.Checked = False
  LightOff.Checked = True
  '
End Sub
Private Sub LightOn_Click()
  '
  FOG_ENABLE = True
  '
  LightOn.Checked = True
  LightOff.Checked = False
  '
End Sub
Private Sub LoadMap_Click()
  '
  ' Show open file dialog
  '
  CommonDialog1.Filter = "Maze File (*.maz)|*.maz"
  CommonDialog1.ShowOpen
  '
  If CommonDialog1.FileName <> "" Then
    '
    ' Load the map
    '
    Call Load_Map(CommonDialog1.FileName)
    '
    ' Set up initial player position
    '
    ViewAngle = PlayerHomeA
    PlayerX = PlayerHomeX
    PlayerY = PlayerHomeY
    '
  End If
  '
End Sub
Private Sub LoadTextures_Click()
  '
  ' Show open file dialog
  '
  Do
    CommonDialog1.Filter = "Texture Files (*.bmp)|*.bmp"
    CommonDialog1.ShowOpen
    '
    If CommonDialog1.FileName <> "" Then
      '
      ' Load the texture file
      '
      Call PictArrayInit1D(TexturesImage, CommonDialog1.FileName, sa2, bmp2, textures())
      '
      msg$ = ""
      '
      If bmp2.bmWidth <> TPAGE_XSIZE Or bmp2.bmHeight <> TPAGE_YSIZE Then
        '
        msg$ = "Invalid bitmap size - must be " & Format$(TPAGE_XSIZE) & "x" & Format$(TPAGE_YSIZE) & ". "
        '
        msg$ = msg$ & "This bitmap is " & Format$(bmp2.bmWidth) & "x" & Format$(bmp2.bmHeight) & "." & vbCrLf
        '
      End If
      '
      If bmp2.bmPlanes <> 1 Or bmp2.bmBitsPixel <> 8 Then
        '
        msg$ = msg$ & "Invalid bitmap - bitmap may only be a single plane, 256 color (8 bit) bitmap."
        '
      End If
      '
      If msg$ <> "" Then
        '
        Beep
        '
        MsgBox msg$, vbCritical, "Error!"
        '
      Else
        '
        Exit Do
        '
      End If
    End If
  Loop
  '
End Sub
Private Sub RenderTimer_Timer()
  '
  ' Disable timer
  '
  Dim speed As Long, ucolumn As Single
  '
  RenderTimer.Enabled = False
  '
  ucolumn = 0
  '
  ' Main rendering loop
  '
  Do
    '
    ' Use entries in key array to determine which way the player
    ' should move (changes suggested by Gary Allen Beebe - 8-24-99)
    '
    ' This actually makes everything smoother, allows multiple keys
    ' to be pressed (move forward and turn at same time!) and allows
    ' the engine to move at full speed - it seems like what I was
    ' doing before only allowed the player to move at half speed or
    ' so, but rendered at full speed - this should help out now!
    '
    ' Thanks, Gary!
    '
    OldPlayerX = PlayerX: OldPlayerY = PlayerY
    '
    If PlyrKeys(vbKeyShift) = True Then
      '
      speed = 2
      '
    Else
      '
      speed = 1
      '
    End If
    '
    If PlyrKeys(vbKeyRight) = True Then
      '
      ViewAngle = ViewAngle + (0.1 * speed)
      '
      If ViewAngle > 6.28318 Then ViewAngle = ViewAngle - 6.28318
      '
      ucolumn = ucolumn + (speed * 5): If ucolumn > 319 Then ucolumn = 0
      '
    End If
    '
    If PlyrKeys(vbKeyLeft) = True Then
      '
      ViewAngle = ViewAngle - (0.1 * speed)
      '
      If ViewAngle < 0 Then ViewAngle = 6.28318 + ViewAngle
      '
      ucolumn = ucolumn - (speed * 5): If ucolumn < 0 Then ucolumn = 319
      '
    End If
    '
    If PlyrKeys(vbKeyUp) = True Then
      '
      PlayerX = PlayerX + Cos(ViewAngle) * (8 * speed)
      PlayerY = PlayerY + Sin(ViewAngle) * (8 * speed)
      '
      If map(PlayerX \ GRID_RES, PlayerY \ GRID_RES) Then
        '
        PlayerX = OldPlayerX: PlayerY = OldPlayerY
        '
      End If
      '
    End If
    '
    If PlyrKeys(vbKeyDown) = True Then
      '
      PlayerX = PlayerX - Cos(ViewAngle) * (8 * speed)
      PlayerY = PlayerY - Sin(ViewAngle) * (8 * speed)
      '
      If map(PlayerX \ GRID_RES, PlayerY \ GRID_RES) Then
        '
        PlayerX = OldPlayerX: PlayerY = OldPlayerY
        '
      End If
      '
    End If
    '
    If PlyrKeys(vbKeyH) = True Then
      '
      ViewAngle = PlayerHomeA
      PlayerX = PlayerHomeX
      PlayerY = PlayerHomeY
      '
    End If
    '
    ' Draw our view of the maze
    '
    Call Draw_Maze(PlayerX, PlayerY, ViewAngle, Int(ucolumn), viewport(), textures())
    '
    ' Display changes to bitmap
    '
    ViewportImage.Refresh
    '
    ' Increase our frame counter
    '
    FPS = FPS + 1
    '
    ' Allow windows time to process
    '
    DoEvents
    '
  Loop Until ALL_STOP
  '
End Sub
Private Sub FPSTimer_Timer()
  '
  ' Display our frame counter and reset (once per second to give FPS)
  '
  RayCastForm.Caption = "True 3D in VB -" + Str$(FPS) + " FPS"
  '
  FPS = 0
  '
End Sub

Private Sub SetColors_Click()
  '
  SetColorForm.Show 1
  '
End Sub
Private Sub UseColor_Click()
  '
  UseColor.Checked = True
  UseTexture.Checked = False
  '
  USE_COLOR = True
  '
End Sub
Private Sub UseTexture_Click()
  '
  UseColor.Checked = False
  UseTexture.Checked = True
  '
  USE_COLOR = False
  '
End Sub
Private Sub ViewportImage_Click()
  '
  msg$ = ""
  msg$ = msg$ + "True 3D in VB" + Chr$(10)
  
  msg$ = msg$ + Chr$(10)
  msg$ = msg$ + "This program is free software; you can redistribute it and/or" + Chr$(10)
  msg$ = msg$ + "modify it under the terms of the GNU General Public License" + Chr$(10)
  msg$ = msg$ + "as published by the Free Software Foundation; either version 2" + Chr$(10)
  msg$ = msg$ + "of the License, or any later version." + Chr$(10)
  msg$ = msg$ + Chr$(10)
  msg$ = msg$ + "This program is distributed in the hope that it will be useful," + Chr$(10)
  msg$ = msg$ + "but WITHOUT ANY WARRANTY; without even the implied warranty of" + Chr$(10)
  msg$ = msg$ + "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" + Chr$(10)
  msg$ = msg$ + "GNU General Public License for more details." + Chr$(10)
  msg$ = msg$ + Chr$(10)

  '
  res% = MsgBox(msg$, 0, "About...")
  '
End Sub
