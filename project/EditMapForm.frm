VERSION 5.00
Begin VB.Form EditMapForm 
   Caption         =   "Edit Map"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   4860
   Icon            =   "EditMapForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MapText 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3132
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3852
   End
End
Attribute VB_Name = "EditMapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
  '
  Dim x As Integer, y As Integer
  '
  MapText.Text = ""
  '
  MapText.Width = EditMapForm.Width - 100
  MapText.height = EditMapForm.height - 300
  '
  MapText.Text = MapText.Text + Trim$(Str$(xsize)) + Chr$(13) + Chr$(10)
  MapText.Text = MapText.Text + Trim$(Str$(ysize)) + Chr$(13) + Chr$(10)
  MapText.Text = MapText.Text + Trim$(Str$((PlayerX - (GRID_RES \ 2)) \ GRID_RES)) + Chr$(13) + Chr$(10)
  MapText.Text = MapText.Text + Trim$(Str$((PlayerY - (GRID_RES \ 2)) \ GRID_RES)) + Chr$(13) + Chr$(10)
  MapText.Text = MapText.Text + Trim$(Str$(ViewAngle)) + Chr$(13) + Chr$(10)
  '
  For y = 0 To ysize - 1
    '
    For x = 0 To xsize - 1
      '
      If map(x, y) = 0 Then
        '
        MapText.Text = MapText.Text + " "
        '
      Else
        '
        MapText.Text = MapText.Text + Trim$(Str$(map(x, y)))
        '
      End If
      '
    Next
    '
    MapText.Text = MapText.Text + Chr$(13) + Chr$(10)
    '
  Next
  '
End Sub
Private Sub Form_Unload(Cancel As Integer)
  '
  Dim x As Integer, y As Integer, t As Integer
  Dim curlin As Integer, ylin As Integer
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
  lin$ = "": PlayerHomeX = 32: PlayerHomeY = 32: PlayerHomeA = 0
  '
  ylin = 0
  '
  For t = 1 To Len(MapText.Text)
    '
    If Mid$(MapText.Text, t, 1) = Chr$(13) Then
      '
      curlin = curlin + 1
      '
      If curlin = 1 Then xsize = Val(lin$)
      If curlin = 2 Then ysize = Val(lin$)
      If curlin = 3 Then PlayerHomeX = (Val(lin$) * GRID_RES) + (GRID_RES \ 2)
      If curlin = 4 Then PlayerHomeY = (Val(lin$) * GRID_RES) + (GRID_RES \ 2)
      If curlin = 5 Then PlayerHomeA = Val(lin$)
      '
      If curlin > 5 Then
        '
        For x = 0 To xsize - 1
          '
          map(x, ylin) = Val(Mid$(lin$, x + 1, 1))
          '
        Next
        '
        ylin = ylin + 1
        '
        If ylin > ysize - 1 Then Exit For
        '
      End If
      '
      lin$ = ""
      '
      t = t + 1
      '
    Else
      '
      lin$ = lin$ + Mid$(MapText.Text, t, 1)
      '
    End If
    '
  Next
  '
End Sub
