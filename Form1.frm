VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Alpha Blender by Tushar Goswami"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picSlider 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   6465
      TabIndex        =   15
      Top             =   0
      Width           =   6525
      Begin ComctlLib.Slider Slider1 
         Height          =   375
         Left            =   2040
         TabIndex        =   16
         Top             =   0
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   661
         _Version        =   327682
         LargeChange     =   10
         Max             =   240
         TickStyle       =   3
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Immediate Blending >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   17
         Top             =   90
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2640
      Top             =   2280
   End
   Begin VB.PictureBox picFrame 
      Align           =   2  'Align Bottom
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   6465
      TabIndex        =   2
      Top             =   2955
      Width           =   6525
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Text            =   "1"
         Top             =   1410
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "<< Animate ! !"
         Height          =   375
         Left            =   4680
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtInterval 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "10"
         Top             =   1410
         Width           =   855
      End
      Begin VB.TextBox txtFadeFrom 
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "enter your filename here"
         Top             =   1770
         Width           =   3735
      End
      Begin VB.TextBox txtFadeIn 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Text            =   "enter your filename here"
         Top             =   2130
         Width           =   3735
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ". . ."
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   4
         Top             =   1770
         Width           =   615
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   ". . ."
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   3
         Top             =   2130
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speed >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   1440
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Interval  >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   10
         Top             =   1455
         Width           =   990
      End
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":0000
         Height          =   1395
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   6105
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fade From >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1815
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fade In >>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   2175
         Width           =   930
      End
   End
   Begin MSComDlg.CommonDialog Cd1 
      Left            =   2760
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "ALL Picture Files (*.jpg, *.bmp, *.gif, *.wmf, *.cur, *.ico) | *.jpg; *.bmp; *.gif; *.wmf; *.cur; *.ico"
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   120
      Picture         =   "Form1.frx":01D7
      ScaleHeight     =   2160
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   600
      Width           =   6255
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2220
      Left            =   2280
      Picture         =   "Form1.frx":0CDC
      ScaleHeight     =   2160
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const AC_SRC_OVER = &H0
' This structure holds the arguments required by Alphablend function to work
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' This is the main API that is blending the pictures
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
' This is a commenly used API function(maybe by me only) which is very helpful to Tranfer ALL the values of a 'Structure'(Type) to a Long variable
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
' Being used by the Timer
Dim Counter As Long
' The BlendFunction 'Structure' is used by the 'AlphaBlend' API function
Dim BF As BLENDFUNCTION
' Actually the AlphaBlend API Function requires a refrence to a "LONG" value containing the values of BlendFunction structure!. This Variale holds the values done in the BlendFunction Structure.
' A Structure (Type) can be converted into a 'Long' value by using the 'RtlMoveMemory' API Function.. See below for its example ;)
Dim lBF As Long

Private Sub cmdBrowse_Click(Index As Integer)
Picture1.AutoSize = True
Picture2.AutoSize = True
Cd1.ShowOpen
Select Case Index
Case 0
    txtFadeFrom = Cd1.FileName
Case 1
    txtFadeIn = Cd1.FileName
End Select
End Sub

Private Sub Command1_Click()

Counter = 0
Picture2.Picture = LoadPicture(txtFadeFrom)

Timer1.Enabled = True
End Sub

Private Sub Form_Load()
txtFadeFrom = App.Path & "\hat0.jpg"
txtFadeIn = App.Path & "\hat1.jpg"
' Don't forget to do this!
    Picture1.AutoRedraw = True
    Picture1.ScaleMode = vbPixels
    
    Picture2.AutoRedraw = True
    Picture2.ScaleMode = vbPixels
End Sub

Private Sub Picture2_Resize()
' Just for resizing ALL the stuff as a new picture file is mentioned
' in the txtFadeFrom textbox :--)
    Me.Height = Picture2.Height + 840 + picFrame.Height + picSlider.Height
If Picture2.Width > 6225 Then
    Me.Width = Picture2.Width + 480
End If
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
End Sub

Private Sub Slider1_Scroll()
Timer1.Enabled = False
Picture2.Picture = LoadPicture(txtFadeFrom)
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Slider1.Value
        .AlphaFormat = 0
    End With
    
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
    Picture2.Refresh

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Counter > 240 Then
    Counter = 240
End If

Counter = Counter + Val(Text1)
    'set the parameters
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = Counter
        .AlphaFormat = 0
    End With
    
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from Picture1 over the picture of Picture2
    AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, lBF
    Picture2.Refresh

If Counter = 240 Then
    Timer1.Enabled = False
End If
End Sub

Private Sub txtFadeFrom_Change()
On Error Resume Next
Picture2.Picture = LoadPicture(txtFadeFrom)
End Sub

Private Sub txtFadeIn_Change()
On Error Resume Next
Picture1.Picture = LoadPicture(txtFadeIn)
End Sub

Private Sub txtInterval_Change()
Timer1.Interval = Val(txtInterval)
End Sub
