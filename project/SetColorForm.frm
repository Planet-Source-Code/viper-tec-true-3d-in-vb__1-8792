VERSION 5.00
Begin VB.Form SetColorForm 
   Caption         =   "Set Colors..."
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3570
   Icon            =   "SetColorForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   3570
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CeilingText 
      Height          =   285
      Left            =   1215
      TabIndex        =   6
      Top             =   1485
      Width           =   2175
   End
   Begin VB.TextBox FloorText 
      Height          =   285
      Left            =   1215
      TabIndex        =   5
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton OKCommand 
      Caption         =   "OK"
      Height          =   420
      Left            =   1080
      TabIndex        =   2
      Top             =   2025
      Width           =   1365
   End
   Begin VB.TextBox FadeToText 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   675
      Width           =   2175
   End
   Begin VB.TextBox ForeColorText 
      Height          =   285
      Left            =   1215
      TabIndex        =   0
      Top             =   270
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ceiling Color :"
      Height          =   285
      Left            =   135
      TabIndex        =   8
      Top             =   1485
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Floor Color :"
      Height          =   285
      Left            =   135
      TabIndex        =   7
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fog Color :"
      Height          =   285
      Left            =   135
      TabIndex        =   4
      Top             =   675
      Width           =   960
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wall Color :"
      Height          =   285
      Left            =   135
      TabIndex        =   3
      Top             =   270
      Width           =   960
   End
End
Attribute VB_Name = "SetColorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  '
  ForeColorText.Text = Str$(FORE_COLOR)
  FadeToText.Text = Str$(FADE_TO_COL)
  FloorText.Text = Str$(FLOOR_COLOR)
  CeilingText.Text = Str$(CEILING_COLOR)
  '
End Sub
Private Sub OKCommand_Click()
  '
  If Val(ForeColorText.Text) >= 0 And Val(ForeColorText.Text) <= 255 Then
    '
    FORE_COLOR = Val(ForeColorText.Text)
    '
  End If
  '
  If Val(FadeToText.Text) >= 0 And Val(FadeToText.Text) <= 255 Then
    '
    FADE_TO_COL = Val(FadeToText.Text)
    '
  End If
  '
  If Val(FloorText.Text) >= 0 And Val(FloorText.Text) <= 255 Then
    '
    FLOOR_COLOR = Val(FloorText.Text)
    '
  End If
  '
  If Val(CeilingText.Text) >= 0 And Val(CeilingText.Text) <= 255 Then
    '
    CEILING_COLOR = Val(CeilingText.Text)
    '
  End If
  '
  Unload Me
  '
End Sub
