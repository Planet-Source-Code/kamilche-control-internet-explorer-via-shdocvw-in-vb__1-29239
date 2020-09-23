VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   Begin VB.Timer tmrMouse 
      Interval        =   100
      Left            =   5340
      Top             =   3870
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   300
      Left            =   13110
      TabIndex        =   28
      Top             =   11205
      Visible         =   0   'False
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   11175
      Left            =   -30
      TabIndex        =   24
      Top             =   -30
      Width           =   13200
      ExtentX         =   23283
      ExtentY         =   19711
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Games"
      Height          =   11160
      Left            =   13185
      TabIndex        =   26
      Top             =   0
      Width           =   2160
      Begin VB.CommandButton cmdTradingPost 
         Caption         =   "Trading Post"
         Height          =   225
         Left            =   105
         TabIndex        =   16
         Top             =   3870
         Width           =   1650
      End
      Begin VB.CommandButton cmdDeckSwabber 
         Caption         =   "Deckswabber"
         Height          =   225
         Left            =   105
         TabIndex        =   14
         Top             =   3420
         Width           =   1650
      End
      Begin VB.CommandButton cmdInventory 
         Caption         =   "Inventory"
         Height          =   225
         Left            =   105
         TabIndex        =   15
         Top             =   3645
         Width           =   1650
      End
      Begin VB.CommandButton cmdHeal 
         Caption         =   "Heal my pets!"
         Height          =   225
         Left            =   105
         TabIndex        =   13
         Top             =   3195
         Width           =   1650
      End
      Begin VB.CommandButton cmdTugOfWar 
         Caption         =   "Tug Of War"
         Height          =   225
         Left            =   105
         TabIndex        =   12
         Top             =   2970
         Width           =   1650
      End
      Begin VB.CommandButton cmdShrine 
         Caption         =   "Shrine"
         Height          =   225
         Left            =   105
         TabIndex        =   10
         Top             =   2520
         Width           =   1650
      End
      Begin VB.CommandButton cmdFruitMachine 
         Caption         =   "Fruit Machine"
         Height          =   225
         Left            =   105
         TabIndex        =   9
         Top             =   2295
         Width           =   1650
      End
      Begin VB.CommandButton cmdTecho 
         Caption         =   "Techo Says"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   2745
         Width           =   1650
      End
      Begin VB.CommandButton cmdWheelofMediocrity 
         Caption         =   "Wheel of Mediocrity"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   2070
         Width           =   1650
      End
      Begin VB.CommandButton cmdWheelOfFortune 
         Caption         =   "Wheel of Fortune"
         Height          =   225
         Left            =   105
         TabIndex        =   7
         Top             =   1845
         Width           =   1650
      End
      Begin VB.CommandButton cmdMoneyTree 
         Caption         =   "Money Tree"
         Height          =   225
         Left            =   105
         TabIndex        =   6
         Top             =   1620
         Width           =   1650
      End
      Begin VB.CommandButton cmdPlaque 
         Caption         =   "Attack of the Plaque"
         Height          =   225
         Left            =   105
         TabIndex        =   5
         Top             =   1395
         Width           =   1650
      End
      Begin VB.CommandButton cmdKacheekSeek 
         Caption         =   "Kacheek Seek"
         Height          =   225
         Left            =   105
         TabIndex        =   4
         Top             =   1170
         Width           =   1650
      End
      Begin VB.CommandButton cmdPoogleSolitaire 
         Caption         =   "Poogle Solitaire"
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   945
         Width           =   1650
      End
      Begin VB.CommandButton cmdBank 
         Caption         =   "Neopets Bank"
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   495
         Width           =   1650
      End
      Begin VB.CommandButton cmdMain 
         Caption         =   "Neopets Home"
         Height          =   225
         Left            =   105
         TabIndex        =   0
         Top             =   270
         Width           =   1650
      End
      Begin VB.CommandButton cmdKiko 
         Caption         =   "Kiko Match"
         Height          =   225
         Left            =   105
         TabIndex        =   2
         Top             =   720
         Width           =   1650
      End
      Begin VB.CommandButton cmdQuit 
         Caption         =   "X"
         Height          =   210
         Left            =   1935
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Width           =   210
      End
      Begin VB.Frame Frame2 
         Caption         =   "Techo Says"
         Height          =   1020
         Left            =   75
         TabIndex        =   42
         Top             =   6705
         Width           =   2010
         Begin VB.CommandButton cmdTechoClear 
            Caption         =   "Clear"
            Height          =   630
            Left            =   1380
            TabIndex        =   23
            Top             =   270
            Width           =   510
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "6"
            Height          =   315
            Index           =   6
            Left            =   960
            TabIndex        =   22
            Top             =   585
            Width           =   405
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "5"
            Height          =   315
            Index           =   5
            Left            =   540
            TabIndex        =   21
            Top             =   585
            Width           =   405
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "4"
            Height          =   315
            Index           =   4
            Left            =   135
            TabIndex        =   20
            Top             =   585
            Width           =   405
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "3"
            Height          =   315
            Index           =   3
            Left            =   960
            TabIndex        =   19
            Top             =   270
            Width           =   405
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "2"
            Height          =   315
            Index           =   2
            Left            =   540
            TabIndex        =   18
            Top             =   270
            Width           =   405
         End
         Begin VB.CommandButton cmdTechoLast 
            Caption         =   "1"
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   17
            Top             =   270
            Width           =   405
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   1425
         Left            =   60
         ScaleHeight     =   1365
         ScaleWidth      =   1995
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   5220
         Width           =   2055
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   11
            Left            =   1485
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   900
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   10
            Left            =   990
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   900
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   9
            Left            =   495
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   900
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   8
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   900
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   7
            Left            =   1485
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   450
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   6
            Left            =   990
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   450
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   5
            Left            =   495
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   450
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   4
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   450
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   3
            Left            =   1485
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   2
            Left            =   990
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   1
            Left            =   495
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
         Begin VB.PictureBox picKiko 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            ForeColor       =   &H80000008&
            Height          =   465
            Index           =   0
            Left            =   0
            ScaleHeight     =   435
            ScaleWidth      =   480
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   0
            Width           =   510
         End
      End
      Begin VB.TextBox Text1 
         Height          =   3300
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   7815
         Width           =   2085
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   25
      Top             =   11145
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19500
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3969
            MinWidth        =   3969
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants
Const MOUSEEVENTF_LEFTDOWN = &H2
Const MOUSEEVENTF_LEFTUP = &H4
Const MOUSEEVENTF_MOVE = &H1
Const KEYEVENTF_KEYUP = &H2
Const VK_LEFT = &H25
Const VK_UP = &H26
Const VK_RIGHT = &H27
Const VK_DOWN = &H28

'Types
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type typeKiko
    X As Long
    Y As Long
    ClickX As Long
    ClickY As Long
    PawX As Long
    PawY As Long
    Color As Long
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

'Declarations
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal uAction As Long) As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'Variables
Private CurrentURL As String
Private Techo() As Long
Private Poogle() As POINTAPI
Private Teeth() As POINTAPI

'--------------------------------------------------------------------------------
' Miscellaneous form routines
'--------------------------------------------------------------------------------
Private Sub Form_Load()
    'Make sure they're 1024x768 mode
    Dim w As Long, h As Long
    w = ScaleX(Screen.Width, vbTwips, vbPixels)
    h = ScaleY(Screen.Height, vbTwips, vbPixels)
    If w < 1024 Or h < 768 Then
        MsgBox "You must be in 1024x768 mode to run this program!" & vbCrLf & _
        "You're currently in " & w & "x" & h, vbCritical
        Unload Me
    End If
End Sub

Private Sub cmdQuit_Click()
    'Quit the program
    Unload Me
End Sub

Private Sub tmrMouse_Timer()
    'Display the XY of the mouse, and the color under the cursor.
    Dim p As POINTAPI
    GetCursorPos p
    StatusBar1.Panels(1).Text = p.X & "," & p.Y & " " & ColorAt(p.X, p.Y)
End Sub

'--------------------------------------------------------------------------------
' Neopets Home
'--------------------------------------------------------------------------------
Private Sub cmdMain_Click()
    NavigateAndWait "http://www.neopets.com"
End Sub

'--------------------------------------------------------------------------------
' Neopets Bank
'--------------------------------------------------------------------------------
Private Sub cmdBank_Click()
    Dim s As String
    'Go to the bank
    NavigateAndWait "http://www.neopets.com/bank.phtml"
    
    'Withdraw any waiting interest
    If ButtonExists("Collect *#* NP of Interest!") Then
        ClickButton "Collect *#* NP of Interest!"
    End If
    
    'Extract the current neopoints on hand
    s = ExtractText("A", "http://www.neopets.com/neopoints.phtml")
    
    'Type it in the first 'amount' field.
    TypeText "amount", Replace(s, ",", "")
    
End Sub

'--------------------------------------------------------------------------------
' Kiko Match
'--------------------------------------------------------------------------------
Private Sub cmdKiko_Click()
    Dim X As Long, Y As Long
    Const w As Long = 214
    Const h As Long = 214
    Const StartX As Long = 98
    Const StartY As Long = 117
    Const PawX As Long = 178
    Const PawY As Long = 196
    Dim Colors(1 To 4, 1 To 3) As typeKiko
    Dim ctr As Long
    
    'Fill the kiko array
    For Y = 1 To 3
        For X = 1 To 4
            With Colors(X, Y)
                .X = X
                .Y = Y
                .ClickX = StartX + ((X - 1) * w)
                .ClickY = StartY + ((Y - 1) * h)
                .PawX = PawX + ((X - 1) * w)
                .PawY = PawY + ((Y - 1) * h)
                .Color = -1
            End With
        Next X
    Next Y
    
    'Navigate to the game
    NavigateAndWait "http://www.neopets.com/games/kikomatch/game.phtml"
    
    'Wait for the 'Play' button, then click it
    StatusBar1.Panels(2).Text = "Waiting for play button..."
    SetCursorPos 659, 624
    Wait 1000
    WaitUntilColorIs 619, 627, 39423
    Wait 1000
    ClickAt 622, 638
    
    'Wait for the screen to fill
    StatusBar1.Panels(2).Text = "Waiting for screen to fill..."
    'ClickAt 457, 482
    WaitUntilColorIs 748, 599, vbBlack
    Wait 1000
    StatusBar1.Panels(2).Text = ""
    
    'Move the cursor out of the way
    SetCursorPos 0, 0
    
    'Go through and grab all the colors
    For Y = 1 To 3
        For X = 1 To 4
            StatusBar1.Panels(2).Text = "Checking " & X & "," & Y
            With Colors(X, Y)
                'Extract the color
                ClickAt .ClickX, .ClickY
                SetCursorPos 0, 0
                Wait 400
                .Color = ColorAt(.ClickX, .ClickY)
                picKiko(ctr).BackColor = .Color
                ctr = ctr + 1
            End With
            If X Mod 2 = 0 Then
                Wait 250
            Else
                Wait 150
            End If
        Next X
    Next Y
    
End Sub

'--------------------------------------------------------------------------------
' Poogle Solitaire
'--------------------------------------------------------------------------------
Private Sub cmdPoogleSolitaire_Click()
    NavigateAndWait "http://www.neopets.com/games/poogle_solitaire/process_poogle_solitaire_new.phtml"
    If WebPageContains("images.neopets.com/pets/angry") Then
        Wait 2000
        NavigateAndWait "http://www.neopets.com/gameroom.phtml"
        Exit Sub
    End If
    PoogleInit
    PoogleMove 29, 17
    PoogleMove 26, 24
    PoogleMove 33, 25
    PoogleMove 18, 30
    PoogleMove 31, 33
    PoogleMove 33, 25
    PoogleMove 6, 18
    PoogleMove 13, 11
    PoogleMove 10, 12
    PoogleMove 27, 13
    PoogleMove 13, 11
    PoogleMove 8, 10
    PoogleMove 1, 9
    PoogleMove 16, 4
    PoogleMove 3, 1
    PoogleMove 1, 9
    PoogleMove 28, 16
    PoogleMove 21, 23
    PoogleMove 24, 22
    PoogleMove 7, 21
    PoogleMove 21, 23
    PoogleMove 10, 8
    PoogleMove 8, 22
    PoogleMove 22, 24
    PoogleMove 24, 26
    PoogleMove 26, 12
    PoogleMove 12, 10
    PoogleMove 17, 15
    PoogleMove 5, 17
    PoogleMove 18, 16
    PoogleMove 15, 17
End Sub

Private Sub PoogleInit()
    ReDim Poogle(1 To 33)
    Poogle(1).X = 335: Poogle(1).Y = 200
    Poogle(2).X = 414: Poogle(2).Y = 200
    Poogle(3).X = 497: Poogle(3).Y = 200
    Poogle(4).X = 335: Poogle(4).Y = 292
    Poogle(5).X = 414: Poogle(5).Y = 292
    Poogle(6).X = 497: Poogle(6).Y = 292
    Poogle(7).X = 174: Poogle(7).Y = 370
    Poogle(8).X = 260: Poogle(8).Y = 370
    Poogle(9).X = 335: Poogle(9).Y = 370
    Poogle(10).X = 414: Poogle(10).Y = 370
    Poogle(11).X = 497: Poogle(11).Y = 370
    Poogle(12).X = 574: Poogle(12).Y = 370
    Poogle(13).X = 645: Poogle(13).Y = 370
    Poogle(14).X = 174: Poogle(14).Y = 456
    Poogle(15).X = 260: Poogle(15).Y = 456
    Poogle(16).X = 335: Poogle(16).Y = 456
    Poogle(17).X = 414: Poogle(17).Y = 456
    Poogle(18).X = 497: Poogle(18).Y = 456
    Poogle(19).X = 574: Poogle(19).Y = 456
    Poogle(20).X = 645: Poogle(20).Y = 456
    Poogle(21).X = 174: Poogle(21).Y = 533
    Poogle(22).X = 260: Poogle(22).Y = 533
    Poogle(23).X = 335: Poogle(23).Y = 533
    Poogle(24).X = 414: Poogle(24).Y = 533
    Poogle(25).X = 497: Poogle(25).Y = 533
    Poogle(26).X = 574: Poogle(26).Y = 533
    Poogle(27).X = 645: Poogle(27).Y = 533
    Poogle(28).X = 335: Poogle(28).Y = 614
    Poogle(29).X = 414: Poogle(29).Y = 614
    Poogle(30).X = 497: Poogle(30).Y = 614
    Poogle(31).X = 335: Poogle(31).Y = 700
    Poogle(32).X = 414: Poogle(32).Y = 700
    Poogle(33).X = 497: Poogle(33).Y = 700
End Sub

Private Sub PoogleMove(ByVal FromSquare As Long, ByVal ToSquare As Long)
    ClickAt Poogle(FromSquare).X, Poogle(FromSquare).Y
    Wait 500
    ClickAndWait Poogle(ToSquare).X, Poogle(ToSquare).Y
    Wait 200
End Sub

'--------------------------------------------------------------------------------
' Kacheek Seek
'--------------------------------------------------------------------------------
Private Sub cmdKacheekSeek_Click()
    
    'Make sure the neopet isn't bored
    NavigateAndWait "http://www.neopets.com/games/hidenseek/0.phtml"
    ClickAndWait 69, 115
    If WebPageContains("images.neopets.com/pets/angry") Then
        Wait 2000
        NavigateAndWait "http://www.neopets.com/gameroom.phtml"
        Exit Sub
    End If
    
    'Play Kacheek Seek in 2 areas.
    SeekIn "happyvalley"
    SeekIn "icecaves"
End Sub

Private Sub SeekIn(ByVal s As String)
    If s = "happyvalley" Then
        NavigateAndWait "http://www.neopets.com/games/hidenseek/0.phtml"
        ClickAndWait 69, 115: GoBack
        ClickAndWait 293, 93: GoBack
        ClickAndWait 167, 268: GoBack
        ClickAndWait 425, 298: GoBack
        ClickAndWait 93, 324: GoBack
    ElseIf s = "icecaves" Then
        NavigateAndWait "http://www.neopets.com/games/hidenseek/1.phtml"
        ClickAndWait 70, 130: GoBack
        ClickAndWait 248, 74: GoBack
        ClickAndWait 351, 107: GoBack
        ClickAndWait 478, 88: GoBack
        ClickAndWait 442, 178: GoBack
        ClickAndWait 325, 157: GoBack
        ClickAndWait 65, 227: GoBack
        ClickAndWait 55, 352: GoBack
        ClickAndWait 225, 349: GoBack
        ClickAndWait 378, 322: GoBack
        ClickAndWait 174, 157: GoBack
    End If
End Sub

'--------------------------------------------------------------------------------
' Attack of the Plaque
'--------------------------------------------------------------------------------
Private Sub cmdPlaque_Click()
    Const Brushing = "Brushing teeth, hit SHIFT to stop."
    Const Paused = "Round finished, brushing paused until you start the next round."
    Const URL = "http://www.neopets.com/games/crest/game.phtml"
    Dim Round As Long, i As Long
    
    Round = 1
    InitTeeth Round
    NavigateAndWait URL
    SetCursorPos 450, 540
    Wait 1000
    WaitUntilColorIs 450, 540, 16768187
    ClickAndWaitForColorChange 450, 540
    Wait 1000
    ClickAt 300, 550
    Wait 2000
    
    'Brush the teeth
    StatusBar1.Panels(2).Text = Brushing
    Do
        For i = LBound(Teeth, 1) To UBound(Teeth, 1)
            BrushTooth i
            If ColorAt(700, 400) = 12412160 Then
                Round = Round + 1
                StatusBar1.Panels(2).Text = Paused
                InitTeeth Round
                Do Until ColorAt(700, 400) <> 12412160
                    DoEvents
                Loop
                StatusBar1.Panels(2).Text = Brushing
            End If
            DoEvents
            If ShiftDown Or CurrentURL <> URL Then
                Exit Do
            End If
        Next i
        If ShiftDown Then
            Exit Do
        End If
    Loop
End Sub

Private Sub BrushTooth(ByVal WhichTooth As Long)
    'Brushes a tooth until it's white, or missing
    Dim FromPt As POINTAPI, ToPt As POINTAPI, c As Long
    Dim i As Long, r As Integer, g As Integer, b As Integer
    FromPt.X = Teeth(WhichTooth).X
    FromPt.Y = Teeth(WhichTooth).Y
    ToPt.X = Teeth(WhichTooth).X
    ToPt.Y = Teeth(WhichTooth).Y + 30
    
    SetCursorPos FromPt.X, FromPt.Y
    Wait 20
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Wait 20
    
    For i = 1 To 15
        c = ColorAt(FromPt.X, FromPt.Y)
        GetRGB c, r, g, b
        If c = vbWhite Then
            Exit For
        ElseIf r > (g + 10) And r > (b + 10) Then
            Exit For
        ElseIf b < 20 And g < 20 Then
            Exit For
        End If
        If ColorAt(700, 400) = 12412160 Then
            Exit For
        End If
        SetCursorPos FromPt.X, FromPt.Y
        Wait 20
        mouse_event MOUSEEVENTF_MOVE, 0, 0, 0, 0
        SetCursorPos ToPt.X, ToPt.Y
        Wait 20
        mouse_event MOUSEEVENTF_MOVE, 0, 0, 0, 0
        If i Mod 5 = 0 Then
            DoEvents
        End If
        'Wait 20
    Next i
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub InitTeeth(ByVal Round As Long)
    If Round = 1 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 643: Teeth(1).Y = 306
        Teeth(2).X = 708: Teeth(2).Y = 275
        Teeth(3).X = 717: Teeth(3).Y = 500
        Teeth(4).X = 772: Teeth(4).Y = 467
    ElseIf Round = 2 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 424: Teeth(1).Y = 377
        Teeth(2).X = 456: Teeth(2).Y = 473
        Teeth(3).X = 643: Teeth(3).Y = 305
        Teeth(4).X = 698: Teeth(4).Y = 282
'        Teeth(5).X = 700: Teeth(5).Y = 498
'        Teeth(6).X = 770: Teeth(6).Y = 476
    ElseIf Round = 3 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 408: Teeth(1).Y = 379
        Teeth(2).X = 531: Teeth(2).Y = 357
        Teeth(3).X = 639: Teeth(3).Y = 308
        Teeth(4).X = 708: Teeth(4).Y = 277
'        Teeth(5).X = 460: Teeth(5).Y = 479
'        Teeth(6).X = 577: Teeth(6).Y = 515
'        Teeth(7).X = 675: Teeth(7).Y = 515
'        Teeth(8).X = 769: Teeth(8).Y = 479
    ElseIf Round = 4 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 443: Teeth(1).Y = 373
        Teeth(2).X = 492: Teeth(2).Y = 373
        Teeth(3).X = 546: Teeth(3).Y = 353
        Teeth(4).X = 597: Teeth(4).Y = 330
'        Teeth(5).X = 668: Teeth(5).Y = 295
'        Teeth(6).X = 716: Teeth(6).Y = 278
'        Teeth(7).X = 439: Teeth(7).Y = 461
'        Teeth(8).X = 470: Teeth(8).Y = 486
'        Teeth(9).X = 527: Teeth(9).Y = 504
'        Teeth(10).X = 598: Teeth(10).Y = 515
'        Teeth(11).X = 679: Teeth(11).Y = 514
'        Teeth(12).X = 769: Teeth(12).Y = 476
    ElseIf Round = 5 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 415: Teeth(1).Y = 372
        Teeth(2).X = 453: Teeth(2).Y = 371
        Teeth(3).X = 490: Teeth(3).Y = 368
        Teeth(4).X = 526: Teeth(4).Y = 356
    ElseIf Round = 6 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 415: Teeth(1).Y = 376
        Teeth(2).X = 453: Teeth(2).Y = 374
        Teeth(3).X = 490: Teeth(3).Y = 373
        Teeth(4).X = 527: Teeth(4).Y = 357
    ElseIf Round = 7 Then
        ReDim Teeth(1 To 4)
        Teeth(1).X = 424: Teeth(1).Y = 373
        Teeth(2).X = 453: Teeth(2).Y = 375
        Teeth(3).X = 484: Teeth(3).Y = 372
        Teeth(4).X = 521: Teeth(4).Y = 362
    End If
End Sub

'--------------------------------------------------------------------------------
' Money Tree
'--------------------------------------------------------------------------------
Private Sub cmdMoneyTree_Click()
    NavigateAndWait "http://www.neopets.com/donations.phtml"
    GrabOne
End Sub

'--------------------------------------------------------------------------------
' Wheel of Fortune
'--------------------------------------------------------------------------------
Private Sub cmdWheelOfFortune_Click()
    NavigateAndWait "http://www.neopets.com/faerieland/wheel.phtml"
    ClickButton "Spin Spin Spin the Wheel of Excitement!"
End Sub

'--------------------------------------------------------------------------------
' Wheel of Mediocrity
'--------------------------------------------------------------------------------
Private Sub cmdWheelofMediocrity_Click()
    NavigateAndWait "http://www.neopets.com/prehistoric/mediocrity.phtml"
    ClickButton "Spin Spin Spin the Wheel of Mediocrity!"
End Sub

'--------------------------------------------------------------------------------
'Fruit Machine
'--------------------------------------------------------------------------------
Private Sub cmdFruitMachine_Click()
    NavigateAndWait "http://www.neopets.com/desert/fruitmachine.phtml"
    ClickButton "Spin The Wheel!!!"
End Sub

'--------------------------------------------------------------------------------
' Shrine
'--------------------------------------------------------------------------------
Private Sub cmdShrine_Click()
    NavigateAndWait "http://www.neopets.com/desert/shrine.phtml"
    ClickButton "Approach the Shrine"
End Sub

'--------------------------------------------------------------------------------
' Techo Says
'--------------------------------------------------------------------------------
Private Sub cmdTecho_Click()
    NavigateAndWait "http://www.neopets.com/games/techosays/game.phtml"
    ReDim Techo(0 To 0)
    StatusBar1.Panels(2).Text = "Click on the LAST note only that Techo played."
End Sub

Private Sub cmdTechoClear_Click()
    ReDim Techo(0 To 0)
End Sub

Private Sub cmdTechoLast_Click(Index As Integer)
    Dim Max As Long, i As Long
    Dim p As POINTAPI
    GetCursorPos p
    'Add the last entry to the list
    Max = UBound(Techo, 1) + 1
    ReDim Preserve Techo(0 To Max)
    Techo(Max) = Index
    'Play the list
    For i = 1 To Max
        Select Case Techo(i)
            Case 1: ClickAt 248, 358
            Case 2: ClickAt 424, 332
            Case 3: ClickAt 588, 359
            Case 4: ClickAt 235, 468
            Case 5: ClickAt 421, 488
            Case 6: ClickAt 607, 471
        End Select
        Wait 200
    Next i
    SetCursorPos p.X, p.Y
End Sub

'--------------------------------------------------------------------------------
' Tug of War
'--------------------------------------------------------------------------------
Private Sub cmdTugOfWar_Click()
    Const URL As String = "http://www.neopets.com/games/tugowar/game.phtml"
    If CurrentURL = URL Then
        StatusBar1.Panels(2).Text = "Tugging, hit 'shift' to stop..."
        WebBrowser1.SetFocus
        DoEvents
        Do
            If ShiftDown = True Or CurrentURL <> URL Then
                Exit Do
            End If
            
            SendScanKey VkKeyScan(Asc("z"))
            SendScanKey VkKeyScan(Asc("x"))
        Loop
        StatusBar1.Panels(2).Text = ""
    Else
        NavigateAndWait "http://www.neopets.com/games/tugowar/game.phtml"
        Wait 1000
        ClickAt 415, 335
        StatusBar1.Panels(2).Text = "Choose 'Horak', password lrlrss"
        Clipboard.Clear
        Clipboard.SetText "lrlrss"
    End If
End Sub

'--------------------------------------------------------------------------------
' Heal My Pets
'--------------------------------------------------------------------------------
Private Sub cmdHeal_Click()
    'Get your pets healed.
    NavigateAndWait "http://www.neopets.com/faerieland/springs.phtml"
    ClickButton "Heal My Pets"
End Sub

'--------------------------------------------------------------------------------
' Deckswabber
'--------------------------------------------------------------------------------
Private Sub cmdDeckSwabber_Click()
    'If you're not there already, go to the game.
    'If you're at the game, navigate the next round.
    Static Round As Long
    If CurrentURL = "http://www.neopets.com/games/deckswabber/game.phtml" Then
        WebBrowser1.SetFocus
        Round = Round + 1
        Select Case Round
            Case 1, 6: WalkTo "dddddddruuuuuuurdddddddruuuuuuurdddddddruuuuuuurdddddddruuuuuuu"
            Case 2, 7: WalkTo "rrrrrrrdddddddllllllluuuuuurrrrrrdddddllllluuuurrrrdddllluu"
            Case 3, 8: WalkTo "rrrrrrrdddddddllllllluuuuuurrrrrrdddddllllluurrdrurululdllu"
            Case 4, 9: WalkTo "uuuuuuullddddddrdlluuuldddluuuuuurrullllddddddrdl"
            Case 5, 10: WalkTo "lldrrrrrulluuurrrdddddddllluulddllluuuuuuurrrdd"
        End Select
    Else
        NavigateAndWait "http://www.neopets.com/games/deckswabber/game.phtml"
        Wait 1000
        ClickAt 760, 428
    End If
End Sub

Private Sub WalkTo(ByVal s As String)
    'Walk along the blocks in Deckswabber according to the pattern passed.
    Dim i As Long, Key As Long
    For i = 1 To Len(s)
        s = Right$(s, Len(s) - 1)
        If s = "d" Then
            Key = VK_DOWN
        ElseIf s = "u" Then
            Key = VK_UP
        ElseIf s = "l" Then
            Key = VK_LEFT
        ElseIf s = "r" Then
            Key = VK_RIGHT
        End If
        SendScanKey Key
        Wait 50
    Next i
End Sub

'--------------------------------------------------------------------------------
' Neopets Inventory
'--------------------------------------------------------------------------------
Private Sub cmdInventory_Click()
    NavigateAndWait "http://www.neopets.com/objects.phtml?type=inventory"
End Sub

'--------------------------------------------------------------------------------
' Trading Post
'--------------------------------------------------------------------------------
Private Sub cmdTradingPost_Click()
    NavigateAndWait "http://www.neopets.com/island/tradingpost.phtml"
End Sub

'--------------------------------------------------------------------------------
' Miscellaneous support routines
'--------------------------------------------------------------------------------
Private Sub NavigateAndWait(ByVal URL As String)
    'Navigate to a URL, and wait for it to fully display.
    Dim StartTime As Long
    StartTime = timeGetTime
    StatusBar1.Panels(3).Text = ""
    Text1.Text = ""
    ProgressBar1.Visible = True
    MousePointer = vbHourglass
    CurrentURL = ""
    WebBrowser1.Navigate URL
    Do Until CurrentURL > ""
        DoEvents
    Loop
    MousePointer = vbDefault
    ProgressBar1.Visible = False
    StatusBar1.Panels(3).Text = "Elapsed: " & Format((timeGetTime - StartTime) / 1000, "#.00") & " seconds"
End Sub

Private Function ShiftDown() As Boolean
    'Returns whether or not the SHIFT key is currently down
    Dim RetVal As Long
    RetVal = GetAsyncKeyState(16) 'SHIFT key
    If (RetVal And 32768) <> 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
End Function

Private Sub ClickAndWait(ByVal X As Long, ByVal Y As Long)
    'Click at a spot on the screen, and wait for the web page to change.
    CurrentURL = ""
    ClickAt X, Y
    Do Until CurrentURL > ""
        DoEvents
    Loop
End Sub

Private Sub GoBack()
    'Hit 'go back' on the browser, and wait for the web page to change.
    CurrentURL = ""
    Wait 20
    WebBrowser1.GoBack
    Do Until CurrentURL > ""
        DoEvents
    Loop
    Wait 20
End Sub

Private Sub ShowIt(ByVal s As String)
    'Show text in the debug textbox.
    Text1.SelStart = Len(Text1.Text)
    Text1.SelText = s & vbCrLf
    Text1.SelStart = Len(Text1.Text)
End Sub

Private Sub ClickAt(ByVal X As Long, ByVal Y As Long)
    'Click at a certain XY position on the screen.
    SetCursorPos X, Y
    Wait 10
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Wait 10
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Public Sub Wait(ByVal milliseconds As Long)
    'Wait the specified number of milliseconds without hanging the program
    Dim StartTime As Long, StopTime As Long
    StartTime = timeGetTime
    StopTime = StartTime + milliseconds
    DoEvents
    Do Until timeGetTime > StopTime
        DoEvents
    Loop
End Sub

Private Function ColorAt(ByVal X As Long, ByVal Y As Long) As Long
    'Return the color at a certain position on the screen.
    Dim TheDC As Long
    TheDC = GetDC(0)
    ColorAt = GetPixel(TheDC, X, Y)
    ReleaseDC 0, TheDC
End Function

Private Function WaitUntilColorIs(ByVal X As Long, ByVal Y As Long, ParamArray Colors() As Variant)
    'Wait until the color at the specific XY location changes.
    Dim TheColor As Long, OrigColor As Variant, FoundIt As Boolean
    TheColor = ColorAt(X, Y)
    FoundIt = False
    Do
        TheColor = ColorAt(X, Y)
        For Each OrigColor In Colors
            If TheColor = OrigColor Then
                FoundIt = True
                Exit For
            End If
        Next
        If FoundIt = True Then
            Exit Do
        End If
        DoEvents
    Loop
End Function

Private Function ClickAndWaitForColorChange(ByVal X As Long, ByVal Y As Long)
    Dim OldColor As Long, NewColor As Long
    OldColor = ColorAt(X, Y)
    ClickAt X, Y
    Do
        NewColor = ColorAt(X, Y)
        If NewColor <> OldColor Then
            Exit Do
        End If
        DoEvents
    Loop
End Function

Private Sub GetRGB(ColorValue As Long, cRed As Integer, cGreen As Integer, cBlue As Integer)
    'Split a color out into its RGB values.
    cRed = ColorValue And &HFF&
    cGreen = (ColorValue And &HFF00&) \ 256
    cBlue = (ColorValue And &HFF0000) \ 65536
End Sub

Private Sub SendScanKey(ByVal TheScanKey As Long)
    'Better than sendkeys - will send arrows, and works in almost all cases
    'whereas SendKeys doesn't work a lot.
    keybd_event TheScanKey, 1, 0, 0
    Wait 50
    keybd_event TheScanKey, 1, KEYEVENTF_KEYUP, 0
    Wait 10
End Sub

'--------------------------------------------------------------------------------
' Miscellaneous web browser routines
'--------------------------------------------------------------------------------

Private Sub GrabOne()
    'Attempts to grab a random item from the money tree.
    Dim c As New HTMLElementCollection
    
    Dim i As Long, Max As Long, ctr As Long
    Dim HTMLElement, All As Collection
    
    'Create a collection to hold the list of items
    Set All = New Collection
    
    'Pick through the web page, looking for donation items
    Max = WebBrowser1.Document.All.length
    For i = 1 To Max
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If HTMLElement.tagName = "A" Then
                If InStr(1, HTMLElement.href, "takedonation.phtml", vbTextCompare) > 0 Then
                    'Add the donation item to the list
                    ctr = ctr + 1
                    All.Add i
                End If
            End If
        End If
    Next i
    
    'Display how many donation items were found.
    StatusBar1.Panels(2).Text = ctr & " items listed, trying to grab one now..."
    
    'Select a random donation
    Max = All.Count
    i = Int((Max * Rnd) + 1)
    
    'Click on the donation, and wait till the web page changes
    CurrentURL = ""
    WebBrowser1.Document.All.Item(All(i)).Click
    Do Until CurrentURL > ""
        DoEvents
    Loop
    
    'Go back to the money tree
    NavigateAndWait "http://www.neopets.com/donations.phtml"
End Sub

Private Function ButtonExists(ByVal s As String) As Boolean
    'Returns whether or not a button exists on the web page
    Dim i As Long, Caption As String, HTMLElement
    
    For i = 1 To WebBrowser1.Document.All.length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If StrComp(HTMLElement.tagName, "INPUT", vbTextCompare) = 0 Then
                If StrComp(HTMLElement.Type, "submit", vbTextCompare) = 0 Then
                    Caption = HTMLElement.Value
                    If (Caption Like s) Or (StrComp(Caption, s, vbTextCompare) = 0) Then
                        ButtonExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

Private Function WebPageContains(ByVal s As String) As Boolean
    'Returns whether or not the web page contains a certain phrase.
    Dim i As Long, HTMLElement
    For i = 1 To WebBrowser1.Document.All.length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If InStr(1, HTMLElement.innerHTML, s, vbTextCompare) > 0 Then
                WebPageContains = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub ClickButton(ByVal s As String)
    'Click a button on the web page
    Dim i As Long, Caption As String, HTMLElement
    On Error GoTo Err_Init
    
    For i = 1 To WebBrowser1.Document.All.length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If StrComp(HTMLElement.tagName, "INPUT", vbTextCompare) = 0 Then
                If StrComp(HTMLElement.Type, "submit", vbTextCompare) = 0 Then
                    Caption = HTMLElement.Value
                    If (Caption Like s) Or (StrComp(Caption, s, vbTextCompare) = 0) Then
                        HTMLElement.Click
                        Exit For
                    End If
                End If
            End If
        End If
    Next i
    
    Exit Sub
Err_Init:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Resume Next
End Sub

Private Function ExtractText(ByVal TagType As String, ByVal TagDescr As String) As String
    'Extract text from a label on the web page
    Dim i As Long, HTMLElement
    On Error GoTo Err_Init
    
    For i = 1 To WebBrowser1.Document.All.length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If StrComp(HTMLElement.tagName, TagType, vbTextCompare) = 0 Then
                If StrComp(HTMLElement.href, TagDescr, vbTextCompare) = 0 Then
                    ExtractText = HTMLElement.outerText
                    Exit For
                End If
            End If
        End If
    Next i
    
    Exit Function
Err_Init:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Resume Next
End Function

Private Sub TypeText(ByVal FieldName As String, ByVal FieldValue As String)
    'Put text in a field on the web page
    Dim i As Long, HTMLElement
    On Error GoTo Err_Init
    
    For i = 1 To WebBrowser1.Document.All.length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
        If Not (HTMLElement Is Nothing) Then
            If StrComp(HTMLElement.tagName, "INPUT", vbTextCompare) = 0 Then
                If StrComp(HTMLElement.Name, FieldName, vbTextCompare) = 0 Then
                    HTMLElement.Value = FieldValue
                    Exit For
                End If
            End If
        End If
    Next i
    
    Exit Sub
Err_Init:
    MsgBox Err.Number & " - " & Err.Description, vbCritical
    Resume Next
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    StatusBar1.Panels(2).Text = "Issuing " & URL
End Sub

Private Sub WebBrowser1_ClientToHostWindow(CX As Long, CY As Long)
    ShowIt "WebBrowser1_ClientToHostWindow"
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
   ' ShowIt "WebBrowser1_CommandStateChange - Command " & Command & ", Enable " & Enable
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    CurrentURL = URL
    StatusBar1.Panels(2).Text = URL
    'ShowIt "WebBrowser1_DocumentComplete"
End Sub

Private Sub WebBrowser1_DownloadBegin()
    ShowIt "WebBrowser1_DownloadBegin"
End Sub

Private Sub WebBrowser1_DownloadComplete()
    ShowIt "WebBrowser1_DownloadComplete"
End Sub

Private Sub WebBrowser1_DragDrop(Source As Control, X As Single, Y As Single)
    ShowIt "WebBrowser1_DragDrop"
End Sub

Private Sub WebBrowser1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    ShowIt "WebBrowser1_DragOver"
End Sub

Private Sub WebBrowser1_FileDownload(Cancel As Boolean)
    ShowIt "WebBrowser1_FileDownload"
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    ShowIt "WebBrowser1_NavigateComplete2"
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
    ShowIt "WebBrowser1_NewWindow2"
End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    If ProgressMax <= 0 Or Progress < 0 Or Progress > ProgressMax Then
        'skip it
    Else
        ProgressBar1.Max = ProgressMax
        ProgressBar1.Value = Progress
    End If
End Sub

Private Sub WebBrowser1_PropertyChange(ByVal szProperty As String)
    ShowIt "WebBrowser1_PropertyChange " & szProperty
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
    'ShowIt "WebBrowser1_StatusTextChange"
    If Text = "Done" Then
        'skip it
    Else
        StatusBar1.Panels(2).Text = Text
    End If
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    ShowIt "WebBrowser1_TitleChange"
End Sub

Private Sub WebBrowser1_Validate(Cancel As Boolean)
    ShowIt "WebBrowser1_Validate"
End Sub

