Attribute VB_Name = "Raycast"

Public Const VIEWPORT_LEFT As Integer = 0
Public Const VIEWPORT_RIGHT As Integer = 320
Public Const VIEWPORT_TOP As Integer = 0
Public Const VIEWPORT_BOT As Integer = 200
Public Const VIEWPORT_HEIGHT As Integer = 200
Public Const VIEWPORT_XCENTER As Integer = 160
Public Const VIEWPORT_YCENTER As Integer = 100
'
Public Const GRID_RES As Long = 64 ' This constant controls wall width
'
Public Const MAX_GRID_X As Long = 64 ' Maximum X dimension of grid
Public Const MAX_GRID_Y As Long = 64 ' Maximum Y dimension of grid
'
Public map(MAX_GRID_X - 1, MAX_GRID_Y - 1) As Integer
Public PlayerHomeA As Single
Public PlayerHomeX As Integer
Public PlayerHomeY As Integer
'
Public xsize As Integer, ysize As Integer
'
Public PlayerX As Integer       ' Player's current X position
Public PlayerY As Integer       ' Player's current Y position
Public ViewAngle As Single      ' Players viewing angle

'
Public viewport() As Byte          ' Viewport bitmap array
Public textures() As Byte          ' Textures bitmap array
Public underlay() As Byte          ' Underlay bitmap array
Public overlay() As Byte           ' Overlay bitmap array
'
' Set up array descriptors
'
Public sa1 As SAFEARRAY1D
Public sa2 As SAFEARRAY1D
Public sa3 As SAFEARRAY1D
Public sa4 As SAFEARRAY1D
'
Public bmp1 As BITMAP
Public bmp2 As BITMAP
Public bmp3 As BITMAP
Public bmp4 As BITMAP
'
' Set up key array (change suggested by Gary Allen Beebe - 8-24-99)
'
Public PlyrKeys(1 To 100) As Boolean
'
' Define Private constants and variables
'
Const WALL_HEIGHT As Integer = 64 ' Controls height of wall
Const VIEWER_HEIGHT As Integer = 32 ' Controls viewer height relative to wall
'
Const VIEWER_DISTANCE As Long = 256
'
Const VDxWH As Long = (VIEWER_DISTANCE * WALL_HEIGHT)
Const VDxVH As Long = (VIEWER_DISTANCE * (WALL_HEIGHT - VIEWER_HEIGHT))
'
' The follow constants are for texture height/width definition. The default
' is for 64x64 pixel textures, but these can be changed to reflect any size
' texture image.
'
Const IMAGE_WIDTH As Integer = 64
Const IMAGE_HEIGHT As Integer = 64
'
Public Const TPAGE_XSIZE As Long = 320 ' These must be public for certain
Public Const TPAGE_YSIZE As Long = 200 ' error detection routines to work
'
Const TILE_XSIZE As Long = TPAGE_XSIZE \ IMAGE_WIDTH
Const TILE_YSIZE As Long = TPAGE_YSIZE \ IMAGE_HEIGHT
'
' These constants are calculated constants used for the grid traversing
' calculations and should not have to be modified.
'
Const MAX_GRID_SIZE As Long = MAX_GRID_X * MAX_GRID_Y
Const GRID_MULT As Long = GRID_RES * (MAX_GRID_SIZE - 1)
'
Dim atn_table(VIEWPORT_RIGHT - VIEWPORT_LEFT) As Single
Sub BuildAtnTable()
  '
  Dim column As Integer
  '
  For column = VIEWPORT_LEFT To VIEWPORT_RIGHT - 1
    '
    atn_table(column - VIEWPORT_LEFT) = Atn((column - (VIEWPORT_RIGHT - VIEWPORT_LEFT) \ 2) / VIEWER_DISTANCE)
    '
  Next
  '
End Sub
Sub Load_Map(datapath As String)
  '
  On Error GoTo fail
  '
  Dim xs As String, ys As String, lin As String
  Dim px As String, py As String, pa As String
  Dim x As Integer, y As Integer
  Dim res As Integer
  '
  ' Clear the map
  '
  For y = 0 To (MAX_GRID_Y - 1)
    '
    For x = 0 To (MAX_GRID_X - 1)
      '
      map(x, y) = 0
      '
    Next
    '
  Next
  '
  Open datapath For Input As 1
  '
  Line Input #1, xs
  Line Input #1, ys
  '
  xsize = Val(xs): ysize = Val(ys)
  '
  If (xsize < 1 Or xsize > MAX_GRID_X) Or (ysize < 1 Or ysize > MAX_GRID_Y) Then
    '
    Beep
    '
    res = MsgBox("Error! - Maze size dimension error!", vbCritical, "Error!")
    '
    Stop
    '
  End If
  '
  Line Input #1, px
  Line Input #1, py
  Line Input #1, pa
  '
  PlayerHomeX = (Val(px) * GRID_RES) + (GRID_RES \ 2)
  PlayerHomeY = (Val(py) * GRID_RES) + (GRID_RES \ 2)
  PlayerHomeA = Val(pa)
  '
  If (PlayerHomeX < (GRID_RES \ 2) Or PlayerHomeX > ((GRID_RES * MAX_GRID_X) - (GRID_RES \ 2))) Or (PlayerHomeY < (GRID_RES \ 2) Or PlayerHomeY > ((GRID_RES * MAX_GRID_Y) - (GRID_RES \ 2))) Then
    '
    Beep
    '
    res = MsgBox("Error! - Player Positioning error!", vbCritical, "Error!")
    '
    Stop
    '
  End If
  '
  For y = 0 To ysize - 1
    '
    Line Input #1, lin
    '
    For x = 0 To xsize - 1
      '
      map(x, y) = Val(Mid$(lin, x + 1, 1))
      '
    Next
    '
  Next
  '
  Close
  '
  Exit Sub
  '
fail:
  '
  If Err = 53 Then
    '
    Beep
    '
    res% = MsgBox("Error! - Maze file not found!", vbCritical, "Error!")
    '
    Stop
    '
  Else
    '
    Resume Next
    '
  End If
  '
  On Error GoTo 0
  '
End Sub
Sub Draw_Maze(xview As Integer, yview As Integer, viewing_angle As Single, ucolumn As Integer, data() As Byte, textures() As Byte)
  '
  Dim sy As Integer, offset As Long
  Dim xd As Single, yd As Single
  Dim bound_x As Integer, bound_y As Integer
  Dim xcross_x As Single, xcross_y As Single
  Dim ycross_x As Single, ycross_y As Single
  Dim xdist As Long, ydist As Long
  Dim xmaze As Integer, ymaze As Integer
  Dim distance As Integer, grid_x As Integer, grid_y As Integer
  '
  Dim column As Integer, tt As Integer
  Dim column_angle As Single, radians As Single
  Dim x2 As Integer, y2 As Integer
  '
  Dim x As Single, y As Single
  Dim xinc As Integer, yinc As Integer
  Dim xdiff As Integer, ydiff As Integer
  Dim slope As Single
  '
  Dim height As Long, bot As Integer, top As Integer
  Dim i As Long, cval As Byte
  '
  Dim tmcolumn As Integer, tyerror As Integer, dheight As Integer
  Dim yratio As Single, tile As Integer, tileptr As Long
  Dim toffset As Integer
  '
  For column = VIEWPORT_LEFT To VIEWPORT_RIGHT - 1
    '
    ucolumn = ucolumn + 1: If ucolumn = 320 Then ucolumn = 0
    '
    column_angle = atn_table(column - VIEWPORT_LEFT)
    radians = viewing_angle + column_angle
    '
    If radians > 6.28318 Then radians = radians - 6.28318
    If radians < 0 Then radians = 6.28318 + radians
    '
    x2 = (GRID_RES * MAX_GRID_X) * Cos(radians) + xview
    y2 = (GRID_RES * MAX_GRID_Y) * Sin(radians) + yview
    '
    x = xview
    y = yview
    '
    xdiff = x2 - xview
    ydiff = y2 - yview
    '
    If xdiff = 0 Then xdiff = 1
    '
    slope = ydiff / xdiff
    '
    If slope = 0 Then slope = 0.0001
    '
    If xdiff > 0 Then
      xinc = GRID_RES
    Else
      xinc = -1
    End If
    '
    If ydiff > 0 Then
      yinc = GRID_RES
    Else
      yinc = -1
    End If
    '
    Do
      '
      grid_x = (x And GRID_MULT) + xinc
      grid_y = (y And GRID_MULT) + yinc
      '
      xcross_x = grid_x
      xcross_y = y + slope * (grid_x - x)
      '
      ycross_x = x + (grid_y - y) / slope
      ycross_y = grid_y
      '
      xd = xcross_x - x
      yd = xcross_y - y
      xdist = Sqr(xd * xd + yd * yd)
      '
      xd = ycross_x - x
      yd = ycross_y - y
      ydist = Sqr(xd * xd + yd * yd)
      '
      If xdist < ydist Then
        '
        xmaze = xcross_x \ GRID_RES
        ymaze = xcross_y \ GRID_RES
        '
        x = xcross_x
        y = xcross_y
        '
        tmcolumn = (y And (IMAGE_WIDTH - 1))
        '
        If x < xview Then tmcolumn = (IMAGE_WIDTH - 1) - tmcolumn
        '
        tile = map(xmaze, ymaze)
        '
        If tile Then Exit Do
        '
      Else
        '
        xmaze = ycross_x \ GRID_RES
        ymaze = ycross_y \ GRID_RES
        '
        x = ycross_x
        y = ycross_y
        '
        tmcolumn = (x And (IMAGE_WIDTH - 1))
        '
        If y > yview Then tmcolumn = (IMAGE_WIDTH - 1) - tmcolumn
        '
        tile = map(xmaze, ymaze)
        '
        If tile Then Exit Do
        '
      End If
      '
    Loop
    '
    xd = x - xview
    yd = y - yview
    '
    distance = Sqr(xd * xd + yd * yd) * Cos(column_angle)
    '
    If distance = 0 Then distance = 1
    '
    height = VDxWH \ distance
    '
    bot = VDxVH \ distance + VIEWPORT_YCENTER
    '
    top = bot - height
    '
    If top < VIEWPORT_TOP Then
      yratio = IMAGE_HEIGHT / height
      toffset = (VIEWPORT_TOP - top) * yratio
      top = VIEWPORT_TOP
    Else
      toffset = 0
    End If
    '
    If bot >= VIEWPORT_BOT Then bot = VIEWPORT_BOT - 1
    '
    Call DrawTexture(top, bot, column, tmcolumn, ucolumn, height, toffset, distance, tile - 1)
    '
  Next
  '
  If Not USE_COLOR Then Call DrawOverlay
  '
End Sub
Sub DrawTexture(top As Integer, bot As Integer, column As Integer, tcolumn As Integer, ucolumn As Integer, height As Long, toffset As Integer, distance As Integer, tile As Integer)
  '
  Dim i As Long
  Dim VideoPointer As Long
  Dim TexturePointer As Long
  Dim tyerror As Long
  Dim cval As Integer, fval As Integer
  '
  If FOG_ENABLE Then
    '
    cval = (distance \ GAMMA) \ 32
    '
    If cval < 0 Then cval = 0
    If cval > 8 Then cval = 8
    '
    ' Draw ground/ceiling with fog fx
    '
    If USE_COLOR Then
      '
      For i = VIEWPORT_TOP To top - 1
        '
        fval = ((i \ GAMMA) \ 8)
        '
        If fval < 0 Then fval = 0
        If fval > 8 Then fval = 8
        '
        viewport(column + i * 320) = trans(FLOOR_COLOR, FADE_TO_COL, fval)
        '
      Next
      '
      For i = bot - 1 To VIEWPORT_BOT - 1
        '
        fval = (((VIEWPORT_BOT - i) \ GAMMA) \ 8)
        '
        If fval < 0 Then fval = 0
        If fval > 8 Then fval = 8
        '
        viewport(column + i * 320) = trans(CEILING_COLOR, FADE_TO_COL, fval)
        '
      Next
      '
    Else
      '
      For i = VIEWPORT_TOP To top - 1
        '
        fval = ((i \ GAMMA) \ 8)
        '
        If fval < 0 Then fval = 0
        If fval > 8 Then fval = 8
        '
        viewport(column + i * 320) = trans(underlay(ucolumn + i * 320), FADE_TO_COL, fval)
        '
      Next
      '
      For i = bot - 1 To VIEWPORT_BOT - 1
        '
        fval = (((VIEWPORT_BOT - i) \ GAMMA) \ 8)
        '
        If fval < 0 Then fval = 0
        If fval > 8 Then fval = 8
        '
        viewport(column + i * 320) = trans(underlay(ucolumn + i * 320), FADE_TO_COL, fval)
        '
      Next
      '
    End If
    '
  Else
    '
    ' Draw ground/ceiling without fog fx
    '
    If USE_COLOR Then
      '
      For i = VIEWPORT_TOP To top - 1
        viewport(column + i * 320) = FLOOR_COLOR
      Next
      '
      For i = bot - 1 To VIEWPORT_BOT - 1
        viewport(column + i * 320) = CEILING_COLOR
      Next
      '
    Else
      '
      For i = VIEWPORT_TOP To top - 1
        viewport(column + i * 320) = underlay(ucolumn + i * 320)
      Next
      '
      For i = bot - 1 To VIEWPORT_BOT - 1
        viewport(column + i * 320) = underlay(ucolumn + i * 320)
      Next
      '
    End If
    '
  End If
  '
  VideoPointer = column + (top * 320)
  '
  tyerror = IMAGE_HEIGHT
  '
  If FOG_ENABLE Then
    '
    If USE_COLOR Then
      For i = 0 To IMAGE_HEIGHT - 1
        While (tyerror >= IMAGE_HEIGHT And VideoPointer < 64000)
          viewport(VideoPointer) = trans(FORE_COLOR, FADE_TO_COL, cval)
          '
          tyerror = tyerror - IMAGE_HEIGHT
          '
          VideoPointer = VideoPointer + 320
        Wend
        '
        tyerror = tyerror + height
      Next
    Else
      TexturePointer = ((tile \ TILE_XSIZE) * TPAGE_XSIZE * IMAGE_HEIGHT) + ((tile Mod TILE_XSIZE) * IMAGE_WIDTH) + (tcolumn + (toffset * TPAGE_XSIZE))
      For i = 0 To IMAGE_HEIGHT - 1
        While (tyerror >= IMAGE_HEIGHT And VideoPointer < 64000)
          viewport(VideoPointer) = trans(textures(TexturePointer), FADE_TO_COL, cval)
          '
          tyerror = tyerror - IMAGE_HEIGHT
          '
          VideoPointer = VideoPointer + 320
        Wend
        '
        tyerror = tyerror + height
        TexturePointer = TexturePointer + TPAGE_XSIZE
      Next
    End If
    '
  Else
    '
    If USE_COLOR Then
      For i = 0 To IMAGE_HEIGHT - 1
        While (tyerror >= IMAGE_HEIGHT And VideoPointer < 64000)
          viewport(VideoPointer) = FORE_COLOR
          '
          tyerror = tyerror - IMAGE_HEIGHT
          '
          VideoPointer = VideoPointer + 320
        Wend
        '
        tyerror = tyerror + height
      Next
    Else
      TexturePointer = ((tile \ TILE_XSIZE) * TPAGE_XSIZE * IMAGE_HEIGHT) + ((tile Mod TILE_XSIZE) * IMAGE_WIDTH) + (tcolumn + (toffset * TPAGE_XSIZE))
      For i = 0 To IMAGE_HEIGHT - 1
        While (tyerror >= IMAGE_HEIGHT And VideoPointer < 64000)
          viewport(VideoPointer) = textures(TexturePointer)
          '
          tyerror = tyerror - IMAGE_HEIGHT
          '
          VideoPointer = VideoPointer + 320
        Wend
        '
        tyerror = tyerror + height
        TexturePointer = TexturePointer + TPAGE_XSIZE
      Next
    End If
    '
  End If
  '
End Sub
Sub DrawOverlay()
  '
  Dim i As Long
  '
  For i = 0 To 63999
    '
    If overlay(i) <> 0 Then
      '
      viewport(i) = overlay(i)
      '
    End If
    '
  Next
  '
End Sub
