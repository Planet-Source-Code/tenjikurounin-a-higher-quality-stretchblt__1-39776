VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SetStretchBltMode API Call"
   ClientHeight    =   7425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picSSetMode 
      AutoRedraw      =   -1  'True
      Height          =   3060
      Left            =   2520
      ScaleHeight     =   3000
      ScaleWidth      =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   3060
   End
   Begin VB.PictureBox picSNoSetMode 
      AutoRedraw      =   -1  'True
      Height          =   3060
      Left            =   2520
      ScaleHeight     =   1910.828
      ScaleMode       =   0  'User
      ScaleWidth      =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   3060
   End
   Begin VB.Frame fraOriginal 
      Caption         =   "Original Image"
      Height          =   2055
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
      Begin VB.PictureBox picOriginal 
         AutoRedraw      =   -1  'True
         Height          =   1560
         Left            =   120
         ScaleHeight     =   1500
         ScaleWidth      =   1500
         TabIndex        =   11
         Top             =   360
         Width           =   1560
      End
   End
   Begin VB.Frame fraStretched 
      Caption         =   "Stretched"
      Height          =   7215
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   3255
      Begin VB.Label lblSNoSetMode 
         Caption         =   "Regular StretchBlt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblSSetMode 
         Caption         =   "With Set Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3840
         Width           =   1215
      End
   End
   Begin VB.PictureBox picRSetMode 
      AutoRedraw      =   -1  'True
      Height          =   1560
      Left            =   360
      ScaleHeight     =   1500
      ScaleWidth      =   1500
      TabIndex        =   4
      Top             =   5640
      Width           =   1560
   End
   Begin VB.PictureBox picRNoSetMode 
      AutoRedraw      =   -1  'True
      Height          =   1560
      Left            =   360
      ScaleHeight     =   955.414
      ScaleMode       =   0  'User
      ScaleWidth      =   1500
      TabIndex        =   5
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Frame fraReduced 
      Caption         =   "Reduced"
      Height          =   4215
      Left            =   240
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
      Begin VB.Label lblRNoSetMode 
         Caption         =   "Regular StretchBlt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblRSetMode 
         Caption         =   "With Set Mode"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API Calls
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal hStretchMode As Long) As Long

Const STRETCHMODE = vbPaletteModeNone   'You can find other modes in the "PaletteModeConstants" section of your Object Browser

Private Sub Form_Load()

    'Load the picture
    picOriginal.Picture = LoadPicture(App.Path & "\kenshin.jpg")

    'Reduce it without SetStretchBltMode
    Call StretchBlt(picRNoSetMode.hdc, 0, 0, 70, 70, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    picRNoSetMode.Refresh
    
    'Reduce it with SetStretchBltMode
    Call SetStretchBltMode(picRSetMode.hdc, STRETCHMODE)
    Call StretchBlt(picRSetMode.hdc, 0, 0, 70, 70, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    picRSetMode.Refresh
    
    'Stretch it without SetStretchBltMode
    Call StretchBlt(picSNoSetMode.hdc, 0, 0, 200, 200, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    picSNoSetMode.Refresh
    
    'Stretch it with SetStretchBltMode
    Call SetStretchBltMode(picSSetMode.hdc, STRETCHMODE)
    Call StretchBlt(picSSetMode.hdc, 0, 0, 200, 200, picOriginal.hdc, 0, 0, 100, 100, vbSrcCopy)
    picSSetMode.Refresh
    
    'Pretty nifty, huh? :)

End Sub
