VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Super Colorpicker v5"
   ClientHeight    =   4095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5325
   Icon            =   "WWWardColorpickerv5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "WWWardColorpickerv5.frx":243A
   ScaleHeight     =   273
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   355
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   210
      Picture         =   "WWWardColorpickerv5.frx":48EFC
      ScaleHeight     =   225
      ScaleWidth      =   390
      TabIndex        =   46
      Top             =   2880
      Width           =   390
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   645
      Picture         =   "WWWardColorpickerv5.frx":493EE
      ScaleHeight     =   225
      ScaleWidth      =   390
      TabIndex        =   45
      Top             =   2880
      Width           =   390
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1995
      Picture         =   "WWWardColorpickerv5.frx":498E0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   43
      Top             =   2640
      Width           =   195
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1995
      Picture         =   "WWWardColorpickerv5.frx":49B2A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   42
      Top             =   2400
      Width           =   195
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1995
      Picture         =   "WWWardColorpickerv5.frx":49D74
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   41
      Top             =   2160
      Width           =   195
   End
   Begin VB.PictureBox Picture8 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4530
      Picture         =   "WWWardColorpickerv5.frx":49FBE
      ScaleHeight     =   195
      ScaleWidth      =   525
      TabIndex        =   40
      Top             =   2925
      Width           =   525
   End
   Begin VB.PictureBox Picture7 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2265
      Picture         =   "WWWardColorpickerv5.frx":4A57C
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   39
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2265
      Picture         =   "WWWardColorpickerv5.frx":4A99A
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   38
      Top             =   2400
      Width           =   375
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2265
      Picture         =   "WWWardColorpickerv5.frx":4ADB8
      ScaleHeight     =   195
      ScaleWidth      =   375
      TabIndex        =   37
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1080
      Picture         =   "WWWardColorpickerv5.frx":4B1D6
      ScaleHeight     =   225
      ScaleWidth      =   1575
      TabIndex        =   36
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox aotd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   4530
      Picture         =   "WWWardColorpickerv5.frx":4C49C
      ScaleHeight     =   120
      ScaleWidth      =   165
      TabIndex        =   35
      ToolTipText     =   "Always on top is off"
      Top             =   75
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.PictureBox aote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   4530
      Picture         =   "WWWardColorpickerv5.frx":4C5FE
      ScaleHeight     =   120
      ScaleWidth      =   165
      TabIndex        =   34
      ToolTipText     =   "Always on top is on"
      Top             =   75
      Width           =   165
   End
   Begin VB.PictureBox GradientR 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   33
      ToolTipText     =   "Drop here the second color for the gradient"
      Top             =   3555
      Width           =   255
   End
   Begin VB.PictureBox GradientL 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   32
      ToolTipText     =   "Drop here the first color for the gradient"
      Top             =   3555
      Width           =   255
   End
   Begin VB.PictureBox Gradient 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   735
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   253
      TabIndex        =   31
      ToolTipText     =   "The gradient, use the 'get pixel' button to grab a color"
      Top             =   3555
      Width           =   3825
   End
   Begin VB.HScrollBar TScroll 
      Height          =   200
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   29
      Top             =   1680
      Width           =   3135
   End
   Begin VB.PictureBox SaveC7 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   2580
      Width           =   255
   End
   Begin VB.PictureBox SaveC6 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   2250
      Width           =   255
   End
   Begin VB.PictureBox SaveC5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   1920
      Width           =   255
   End
   Begin VB.PictureBox SaveC4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   1590
      Width           =   255
   End
   Begin VB.PictureBox SaveC3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   24
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   1260
      Width           =   255
   End
   Begin VB.PictureBox SaveC2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   23
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   930
      Width           =   255
   End
   Begin VB.PictureBox SaveC1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4680
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      ToolTipText     =   "Doubleclick to set as current color"
      Top             =   600
      Width           =   255
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   4800
      Picture         =   "WWWardColorpickerv5.frx":4C760
      ScaleHeight     =   105
      ScaleWidth      =   120
      TabIndex        =   20
      ToolTipText     =   "Minimize"
      Top             =   80
      Width           =   120
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   5040
      Picture         =   "WWWardColorpickerv5.frx":4C84A
      ScaleHeight     =   105
      ScaleWidth      =   120
      TabIndex        =   19
      ToolTipText     =   "Close"
      Top             =   80
      Width           =   120
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   1680
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      Picture         =   "WWWardColorpickerv5.frx":4C934
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   240
      Width           =   135
   End
   Begin VB.TextBox VbVal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   2640
      Width           =   1080
   End
   Begin VB.TextBox HexVal 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   195
      Left            =   840
      TabIndex        =   10
      Top             =   2400
      Width           =   1080
   End
   Begin VB.TextBox Btxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Text            =   "0"
      Top             =   2160
      Width           =   360
   End
   Begin VB.TextBox Gtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Text            =   "0"
      Top             =   2160
      Width           =   360
   End
   Begin VB.TextBox Rtxt 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Text            =   "0"
      Top             =   2160
      Width           =   360
   End
   Begin VB.PictureBox NewPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   2880
      ScaleHeight     =   705
      ScaleWidth      =   1455
      TabIndex        =   6
      ToolTipText     =   "Drag this color to one of the saveslots or to the gradient"
      Top             =   2280
      Width           =   1455
   End
   Begin VB.PictureBox BPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   1320
      Width           =   255
   End
   Begin VB.PictureBox GPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   990
      Width           =   255
   End
   Begin VB.PictureBox RPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   4080
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   630
      Width           =   255
   End
   Begin VB.HScrollBar BScroll 
      Height          =   200
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   2
      Top             =   1350
      Width           =   3135
   End
   Begin VB.HScrollBar GScroll 
      Height          =   200
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   1
      Top             =   990
      Width           =   3135
   End
   Begin VB.HScrollBar RScroll 
      Height          =   200
      LargeChange     =   10
      Left            =   840
      Max             =   255
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   2760
      ScaleHeight     =   930
      ScaleWidth      =   1665
      TabIndex        =   44
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   30
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Super Colorpicker v5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   255
      Left            =   720
      TabIndex        =   21
      ToolTipText     =   "Super Colorpicker v5 :: Created By Wwward :: w_w_w_ard@hotmail.com"
      Top             =   15
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "VB:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "HEX:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "RGB:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Blue:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Green:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Red:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'_/
'_/ Super Colorpicker V4 Created By Wwward
'_/ November-December 2003
'_/ Email me if you like it : w_w_w_ard@hotmail.com
'_/
'_/ At this moment there is no helpfile included
'_/ because I think it speaks for itselves,
'_/ most of the help stands in the tooltips in the program
'_/ so if you have troubles using it, check the tooltips
'_/
'_/ If you use it or change it, give me the credits
'_/
'_/ You will see, the code is well commented ;-) NOT
'_/
'_/ Thanks to: - N.Gopi Krishna: for the code of
'_/              getting a screenpixel color
'_/            - Serge Lachapelle: for the skinning code
'_/            - Adam Jacob Muller: for the always on top
'_/              code
'_/
'_/
'_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

Dim rgbvalue As Long
Dim pt As POINTAPI
Dim dragcolorr
Dim dragcolorg
Dim dragcolorb
Dim TotalChange As String
Dim MiddleVal As Long
Dim Rcolor1
Dim Gcolor1
Dim Bcolor1
    
Dim Rcolor2
Dim Gcolor2
Dim Bcolor2
Dim activetxt As String

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXPos As Long, ByVal nYPos As Long) As Long
Private Declare Function CreateDC Lib "gdi32.dll" Alias "CreateDCA" (ByVal lpszDriver As String, ByVal lpszDevice As String, ByVal lpszOutput As Long, lpInitData As Any) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal HWND As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Sub SetPrev(RVal, GVal, BVal)
    NewPrev.BackColor = RGB(RVal, GVal, BVal)
End Sub
Sub SetBars(RVal, GVal, BVal)
    RScroll.Value = RVal
    GScroll.Value = GVal
    BScroll.Value = BVal
End Sub
Sub SetHex(RVal, GVal, BVal)
    Rhex = Hex(RVal)
    If Len(CStr(Rhex)) < 2 Then Rhex = "0" & Rhex
    Ghex = Hex(GVal)
    If Len(CStr(Ghex)) < 2 Then Ghex = "0" & Ghex
    Bhex = Hex(BVal)
    If Len(CStr(Bhex)) < 2 Then Bhex = "0" & Bhex
    Vbstr = "&H" & Bhex & Ghex & Rhex & "&"
    VbVal.Text = Vbstr
End Sub
Sub SetVb(RVal, GVal, BVal)
    Rhex = Hex(RVal)
    If Len(CStr(Rhex)) < 2 Then Rhex = "0" & Rhex
    Ghex = Hex(GVal)
    If Len(CStr(Ghex)) < 2 Then Ghex = "0" & Ghex
    Bhex = Hex(BVal)
    If Len(CStr(Bhex)) < 2 Then Bhex = "0" & Bhex
    hexstr = "#" & Rhex & Ghex & Bhex
    HexVal.Text = hexstr
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MinTimer.Interval = 0
    MinTimer.Enabled = False
End Sub

Private Sub aotd_Click()
    AlwaysOnTop Form1, True
    aotd.Visible = False
    aote.Visible = True
End Sub

Private Sub aote_Click()
    AlwaysOnTop Form1, False
    aote.Visible = False
    aotd.Visible = True
End Sub

'Private Sub Check1_Click()
'    If Check1.Value = 1 Then
'        AlwaysOnTop Form1, True
'    Else
'        AlwaysOnTop Form1, False
'    End If
'End Sub

Private Sub Form_Load()
If GetSetting("ColorpickerV5", "Colors", "Color1") <> "" Then
    SaveC1.BackColor = GetSetting("ColorpickerV5", "Colors", "Color1")
End If
If GetSetting("ColorpickerV5", "Colors", "Color2") <> "" Then
    SaveC2.BackColor = GetSetting("ColorpickerV5", "Colors", "Color2")
End If
If GetSetting("ColorpickerV5", "Colors", "Color3") <> "" Then
    SaveC3.BackColor = GetSetting("ColorpickerV5", "Colors", "Color3")
End If
If GetSetting("ColorpickerV5", "Colors", "Color4") <> "" Then
    SaveC4.BackColor = GetSetting("ColorpickerV5", "Colors", "Color4")
End If
If GetSetting("ColorpickerV5", "Colors", "Color5") <> "" Then
    SaveC5.BackColor = GetSetting("ColorpickerV5", "Colors", "Color5")
End If
If GetSetting("ColorpickerV5", "Colors", "Color6") <> "" Then
    SaveC6.BackColor = GetSetting("ColorpickerV5", "Colors", "Color6")
End If
If GetSetting("ColorpickerV5", "Colors", "Color7") <> "" Then
    SaveC7.BackColor = GetSetting("ColorpickerV5", "Colors", "Color7")
End If
    AlwaysOnTop Form1, True
    Rcolor1 = 0
    Gcolor1 = 0
    Bcolor1 = 0
    
    Rcolor2 = 0
    Gcolor2 = 0
    Bcolor2 = 0
    TotalChange = "yes"
    RPrev.BackColor = RGB(RScroll.Value, 0, 0)
    GPrev.BackColor = RGB(0, GScroll.Value, 0)
    BPrev.BackColor = RGB(0, 0, BScroll.Value)
    Call SetPrev(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetHex(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetVb(RScroll.Value, GScroll.Value, BScroll.Value)
    
    Picture1.AutoSize = True
    ChangeMask Me, Picture1
    Picture1.Visible = False
    Timer1.Enabled = False
End Sub
Function GetRed(ByVal lColor As Long) As Long
    GetRed = lColor Mod 256
End Function

Function GetGreen(ByVal lColor As Long) As Long
    GetGreen = (lColor \ &H100) Mod 256
End Function

Function GetBlue(ByVal lColor As Long) As Long
    GetBlue = (lColor \ &H10000) Mod 256
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  FormDrag Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "ColorpickerV5", "Colors", "Color1", SaveC1.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color2", SaveC2.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color3", SaveC3.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color4", SaveC4.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color5", SaveC5.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color6", SaveC6.BackColor
    SaveSetting "ColorpickerV5", "Colors", "Color7", SaveC7.BackColor
End Sub

Private Sub Gradient_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub

Private Sub GradientL_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub

Private Sub GradientR_Click()
    Timer1.Interval = 0
    Timer1.Enabled = False
End Sub

Private Sub Label7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub NewPrev_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    dragcolorr = GetRed(NewPrev.BackColor)
    dragcolorg = GetGreen(NewPrev.BackColor)
    dragcolorb = GetBlue((NewPrev.BackColor))
End Sub

Private Sub Picture10_Click()
    Clipboard.Clear
    Clipboard.SetText HexVal.Text
End Sub

Private Sub Picture11_Click()
    Clipboard.Clear
    Clipboard.SetText VbVal.Text
End Sub
Function GetNearestWebColor(colcheck)
colcheck2 = colcheck - 26

If colcheck2 < 0 Then
    GetNearestWebColor = 0
Else
    If colcheck2 < 51 Then
        GetNearestWebColor = 51
    Else
        If colcheck2 < 102 Then
            GetNearestWebColor = 102
        Else
            If colcheck2 < 153 Then
                GetNearestWebColor = 153
            Else
                If colcheck2 < 204 Then
                    GetNearestWebColor = 204
                Else
                    If colcheck2 < 255 Then
                        GetNearestWebColor = 255
                    End If
                End If
            End If
        End If
    End If
End If
End Function

Private Sub Picture13_Click()
    Call SetBars(GetNearestWebColor(GetRed(NewPrev.BackColor)), GetNearestWebColor(GetGreen(NewPrev.BackColor)), GetNearestWebColor(GetBlue(NewPrev.BackColor)))
End Sub

Private Sub Picture14_Click()
    Call SetBars((255 - GetRed(NewPrev.BackColor)), (255 - GetGreen(NewPrev.BackColor)), (255 - GetBlue(NewPrev.BackColor)))
End Sub

Private Sub Picture2_Click()
    Unload Me
End Sub

Private Sub Picture2_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    MsgBox ("ok")
End Sub

Private Sub Picture3_Click()
    WindowState = vbMinimized
End Sub

Private Sub Picture4_Click()
    activetxt = GetActiveWindow
    Interval = 100
    If Timer1.Interval = 0 Then
        Timer1.Enabled = True
        Timer1.Interval = Interval
    Else
        Timer1.Interval = 0
        Timer1.Enabled = False
    End If
End Sub

Private Sub Picture5_Click()
        '---
    If Trim(Rtxt) = "" Then Rtxt.Text = 0
    If IsNumeric(Rtxt) Then
        If Rtxt < 0 Then Rtxt.Text = 0
        If Rtxt > 255 Then Rtxt.Text = 255
        RScroll.Value = Rtxt
    Else
        Rtxt.Text = 0
    End If
    If Trim(Gtxt) = "" Then Gtxt.Text = 0
    If IsNumeric(Gtxt) Then
        If Gtxt < 0 Then Gtxt.Text = 0
        If Gtxt > 255 Then Gtxt.Text = 255
        GScroll.Value = Gtxt
    Else
        Gtxt.Text = 0
    End If
        If Trim(Btxt) = "" Then Btxt.Text = 0
    If IsNumeric(Btxt) Then
        If Btxt < 0 Then Btxt.Text = 0
        If Btxt > 255 Then Btxt.Text = 255
        BScroll.Value = Btxt
    Else
        Btxt.Text = 0
    End If
End Sub

Private Sub Picture6_Click()
    On Error GoTo Errorhand
    hexstr = HexVal.Text
    firstchar = Mid(hexstr, 1, 1)
    If firstchar = "#" Then
        hexstr = Mid(hexstr, 2, 6)
    End If
    If Len(hexstr) = 6 Then
        RVal = CInt("&H" & Mid(hexstr, 1, 2))
        GVal = CInt("&H" & Mid(hexstr, 3, 2))
        BVal = CInt("&H" & Mid(hexstr, 5, 2))
        Call SetBars(RVal, GVal, BVal)
    Else
        HexVal.Text = "#000000"
        Call SetBars(0, 0, 0)
    End If
    
Errorhand:
    If Err.Number <> 0 Then
        HexVal.Text = "#000000"
        Call SetBars(0, 0, 0)
    End If
End Sub

Private Sub Picture7_Click()
    On Error GoTo Errorhand
    Vbstr = VbVal.Text
    firstchars = Mid(Vbstr, 1, 2)
    If firstchars = "&H" Then
        Vbstr = Mid(Vbstr, 3, 6)
    End If
    If Len(Vbstr) = 6 Then
        RVal = CInt("&H" & Mid(Vbstr, 5, 2))
        GVal = CInt("&H" & Mid(Vbstr, 3, 2))
        BVal = CInt("&H" & Mid(Vbstr, 1, 2))
        Call SetBars(RVal, GVal, BVal)
    Else
        VbVal.Text = "&H000000&"
        Call SetBars(0, 0, 0)
    End If
    
Errorhand:
    If Err.Number <> 0 Then
        VbVal.Text = "&H000000&"
        MsgBox Err.Description
        Call SetBars(0, 0, 0)
    End If
End Sub

Private Sub Picture8_Click()
    SaveC1.BackColor = vbBlack
    SaveC2.BackColor = vbBlack
    SaveC3.BackColor = vbBlack
    SaveC4.BackColor = vbBlack
    SaveC5.BackColor = vbBlack
    SaveC6.BackColor = vbBlack
    SaveC7.BackColor = vbBlack
End Sub

Private Sub Picture9_Click()
    Clipboard.Clear
    Clipboard.SetText Rtxt.Text & " " & Gtxt.Text & " " & Btxt.Text
End Sub

Private Sub RScroll_Scroll()
    Call RScroll_Change
End Sub
Private Sub RScroll_Change()
    TotalChange = "no"
    RPrev.BackColor = RGB(RScroll.Value, 0, 0)
    Call SetPrev(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetHex(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetVb(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetTotalBar
    Rtxt.Text = RScroll.Value
    TotalChange = "yes"
End Sub
Private Sub GScroll_Scroll()
    Call GScroll_Change
End Sub
Private Sub GScroll_Change()
    TotalChange = "no"
    GPrev.BackColor = RGB(0, GScroll.Value, 0)
    Call SetPrev(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetHex(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetVb(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetTotalBar
    Gtxt.Text = GScroll.Value
    TotalChange = "yes"
End Sub
Private Sub BScroll_Scroll()
    Call BScroll_Change
End Sub
Private Sub BScroll_Change()
    TotalChange = "no"
    BPrev.BackColor = RGB(0, 0, BScroll.Value)
    Call SetPrev(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetHex(RScroll.Value, GScroll.Value, BScroll.Value)
    Call SetVb(RScroll.Value, GScroll.Value, BScroll.Value)
        Call SetTotalBar
    Btxt.Text = BScroll.Value
    TotalChange = "yes"
End Sub

Private Sub SkinLabel5_DragDrop(Source As Control, X As Single, Y As Single)

End Sub


Private Sub Rtxt_GotFocus()
    Rtxt.SelStart = 0
    Rtxt.SelLength = Len(Rtxt.Text)
End Sub
Private Sub Gtxt_GotFocus()
    Gtxt.SelStart = 0
    Gtxt.SelLength = Len(Gtxt.Text)
End Sub
Private Sub Btxt_GotFocus()
    Btxt.SelStart = 0
    Btxt.SelLength = Len(Btxt.Text)
End Sub
Private Sub HexVal_GotFocus()
    HexVal.SelStart = 0
    HexVal.SelLength = Len(HexVal.Text)
End Sub

Public Sub SetBarsFromLongColor(longcolor)
    ThisR = longcolor Mod &H100
    longcolor = longcolor \ &H100
    ThisG = longcolor Mod &H100
    longcolor = longcolor \ &H100
    ThisB = longcolor Mod &H100
    Call SetBars(ThisR, ThisG, ThisB)
End Sub

Private Sub SaveC1_DblClick()
    SetBarsFromLongColor (SaveC1.BackColor)
End Sub
Private Sub SaveC2_DblClick()
    SetBarsFromLongColor (SaveC2.BackColor)
End Sub
Private Sub SaveC3_DblClick()
    SetBarsFromLongColor (SaveC3.BackColor)
End Sub
Private Sub SaveC4_DblClick()
    SetBarsFromLongColor (SaveC4.BackColor)
End Sub
Private Sub SaveC5_DblClick()
    SetBarsFromLongColor (SaveC5.BackColor)
End Sub
Private Sub SaveC6_DblClick()
    SetBarsFromLongColor (SaveC6.BackColor)
End Sub
Private Sub SaveC7_DblClick()
    SetBarsFromLongColor (SaveC7.BackColor)
End Sub

Private Sub SaveC1_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC1.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC2_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC2.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC3_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC3.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC4_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC4.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC5_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC5.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC6_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC6.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
Private Sub SaveC7_DragDrop(Source As Control, X As Single, Y As Single)
    SaveC7.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
End Sub
'Dragged colors for gradient
Private Sub GradientL_DragDrop(Source As Control, X As Single, Y As Single)
    GradientL.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
    Rcolor1 = dragcolorr
    Gcolor1 = dragcolorg
    Bcolor1 = dragcolorb
    Call SetGradient
End Sub
Private Sub GradientR_DragDrop(Source As Control, X As Single, Y As Single)
    GradientR.BackColor = RGB(dragcolorr, dragcolorg, dragcolorb)
    Rcolor2 = dragcolorr
    Gcolor2 = dragcolorg
    Bcolor2 = dragcolorb
    Call SetGradient
End Sub
'End dragged colors for gradient
Private Sub TScroll_Scroll()
    Call TScroll_Change
End Sub
Private Sub TScroll_Change()
    If TotalChange = "yes" Then
        changeplusmin = CLng(TScroll.Value - MiddleVal)
        'MsgBox changeplusmin
        Call DarkerLighter(Val(changeplusmin))
    End If
End Sub
Public Sub DarkerLighter(plusval As Long)
    If TotalChange = "yes" Then
        Rv = RScroll.Value
        Gv = GScroll.Value
        Bv = BScroll.Value
        Dim minval As Long
        Dim minplusval As Long
        
        minval = Rv
        If Gv < minval Then
            minval = Gv
        End If
        If Bv < minval Then
            minval = Bv
        End If
        
        maxval = Rv
        If Gv > maxval Then
            maxval = Gv
        End If
        If Bv > maxval Then
            maxval = Bv
        End If
        
        maxplus = 255 - maxval
        minplusval = plusval
        
        minval = 0 - minval
        
        If minplusval < minval Then
            minplusval = minval
        End If
        
        If minplusval > maxplus Then
            minplusval = maxplus
        End If
        
        nRv = Rv + minplusval
        nGv = Gv + minplusval
        nBv = Bv + minplusval
        Call SetBars(nRv, nGv, nBv)
        Call SetTotalBar
    End If
End Sub

Private Sub VbVal_GotFocus()
    VbVal.SelStart = 0
    VbVal.SelLength = Len(VbVal.Text)
End Sub
Private Sub Timer1_Timer()
Dim stdout  ' Declares standard output. Defaults to monitor

activewindow = GetActiveWindow
If activewindow <> activetxt Then
    Timer1.Interval = 0
    Timer1.Enabled = False
    Form1.SetFocus
Else
Dim HexEquivalent As String
HexEquivalent = "&HFFFFFFFF"
On Error GoTo 10
GetCursorPos pt
    rgbvalue = GetPixel(GetDC(stdout), pt.X, pt.Y)
    HexEquivalent = Hex(rgbvalue)
    sColor = Right("000000" & HexEquivalent, 6)
        R = Val("&h" & Right(sColor, 2))
        G = Val("&h" & Mid(sColor, 3, 2))
        B = Val("&h" & Left(sColor, 2))
    Call SetBars(R, G, B)
End If
    Exit Sub
10: Exit Sub
End Sub

Private Sub SetTotalBar()
    Rval1 = RScroll.Value
    Gval1 = GScroll.Value
    Bval1 = BScroll.Value
    
    If ((Gval1 <= Rval1 And Bval1 >= Rval1) Or (Gval1 >= Rval1 And Bval1 <= Rval1)) Then
        MiddleVal = Rval1
    End If
    If ((Rval1 <= Gval1 And Bval1 >= Gval1) Or (Rval1 >= Gval1 And Bval1 <= Gval1)) Then
        MiddleVal = Gval1
    End If
    If ((Rval1 <= Bval1 And Gval1 >= Bval1) Or (Rval1 >= Bval1 And Gval1 <= Bval1)) Then
        MiddleVal = Bval1
    End If
    TScroll.Value = MiddleVal
End Sub
Private Sub SetGradient()
    
    gh = Gradient.Height
    gw = Gradient.Width

    Dim Rpl As Double
    Dim Gpl As Double
    Dim Bpl As Double
    Rpl = (Rcolor2 - Rcolor1) / gw
    Gpl = (Gcolor2 - Gcolor1) / gw
    Bpl = (Bcolor2 - Bcolor1) / gw

    tel = 0
    Do While tel < gw
        Rplcol = Round(Rcolor1 + (Rpl * (tel + 1)))
        Gplcol = Round(Gcolor1 + (Gpl * (tel + 1)))
        Bplcol = Round(Bcolor1 + (Bpl * (tel + 1)))
        Gradient.Line (tel, 0)-(tel, gh), RGB(Rplcol, Gplcol, Bplcol)
        tel = tel + 1
    Loop
End Sub
 
Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)


    If SetOnTop Then
        lFlag = HWND_TOPMOST
    Else
        lFlag = HWND_NOTOPMOST
    End If
    SetWindowPos myfrm.HWND, lFlag, _
    myfrm.Left / Screen.TwipsPerPixelX, _
    myfrm.Top / Screen.TwipsPerPixelY, _
    myfrm.Width / Screen.TwipsPerPixelX, _
    myfrm.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub


