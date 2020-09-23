VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00BDFCFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chinese Chess"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MingLiU"
      Size            =   8.25
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   31
      Left            =   5760
      Picture         =   "frmMain.frx":0442
      Tag             =   "castle|89"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   30
      Left            =   5040
      Picture         =   "frmMain.frx":0875
      Tag             =   "knight|79"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   29
      Left            =   4320
      Picture         =   "frmMain.frx":0CB4
      Tag             =   "bishop|69"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   28
      Left            =   3600
      Picture         =   "frmMain.frx":10FC
      Tag             =   "scholar|59"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   27
      Left            =   2880
      Picture         =   "frmMain.frx":1513
      Tag             =   "king|49"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   26
      Left            =   2160
      Picture         =   "frmMain.frx":1969
      Tag             =   "scholar|39"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   25
      Left            =   1440
      Picture         =   "frmMain.frx":1D80
      Tag             =   "bishop|29"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   24
      Left            =   720
      Picture         =   "frmMain.frx":21C8
      Tag             =   "knight|19"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   23
      Left            =   0
      Picture         =   "frmMain.frx":2607
      Tag             =   "castle|09"
      Top             =   0
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   22
      Left            =   5040
      Picture         =   "frmMain.frx":2A3A
      Tag             =   "rocket|77"
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   21
      Left            =   720
      Picture         =   "frmMain.frx":2E8D
      Tag             =   "rocket|17"
      Top             =   1440
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   20
      Left            =   5760
      Picture         =   "frmMain.frx":32E0
      Tag             =   "pawn|86"
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   19
      Left            =   4320
      Picture         =   "frmMain.frx":3712
      Tag             =   "pawn|66"
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   18
      Left            =   2880
      Picture         =   "frmMain.frx":3B44
      Tag             =   "pawn|46"
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   17
      Left            =   1440
      Picture         =   "frmMain.frx":3F76
      Tag             =   "pawn|26"
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   16
      Left            =   0
      Picture         =   "frmMain.frx":43A8
      Tag             =   "pawn|06"
      Top             =   2160
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   15
      Left            =   5760
      Picture         =   "frmMain.frx":47DA
      Tag             =   "pawn|83"
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   14
      Left            =   4320
      Picture         =   "frmMain.frx":4C0C
      Tag             =   "pawn|63"
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   13
      Left            =   2880
      Picture         =   "frmMain.frx":503E
      Tag             =   "pawn|43"
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   12
      Left            =   1440
      Picture         =   "frmMain.frx":5470
      Tag             =   "pawn|23"
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   11
      Left            =   0
      Picture         =   "frmMain.frx":58A2
      Tag             =   "pawn|03"
      Top             =   4320
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   10
      Left            =   5040
      Picture         =   "frmMain.frx":5CD4
      Tag             =   "rocket|72"
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   9
      Left            =   720
      Picture         =   "frmMain.frx":6129
      Tag             =   "rocket|12"
      Top             =   5040
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   8
      Left            =   5760
      Picture         =   "frmMain.frx":657E
      Tag             =   "castle|80"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   7
      Left            =   5040
      Picture         =   "frmMain.frx":69B1
      Tag             =   "knight|70"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   6
      Left            =   4320
      Picture         =   "frmMain.frx":6DF0
      Tag             =   "bishop|60"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   5
      Left            =   3600
      Picture         =   "frmMain.frx":7233
      Tag             =   "scholar|50"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   4
      Left            =   2880
      Picture         =   "frmMain.frx":7668
      Tag             =   "king|40"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   3
      Left            =   2160
      Picture         =   "frmMain.frx":7ABF
      Tag             =   "scholar|30"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   2
      Left            =   1440
      Picture         =   "frmMain.frx":7EF4
      Tag             =   "bishop|20"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   1
      Left            =   720
      Picture         =   "frmMain.frx":8337
      Tag             =   "knight|10"
      Top             =   6480
      Width           =   600
   End
   Begin VB.Image imgPieces 
      Height          =   615
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":8776
      Tag             =   "castle|00"
      Top             =   6480
      Width           =   600
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   3000
      Top             =   3360
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   65
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   51
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   76
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   67
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   55
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   56
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   89
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   88
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   87
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   86
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   85
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   84
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   83
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   82
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   81
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   80
      Left            =   5760
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   79
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   78
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   77
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   75
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   74
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   73
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   72
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   71
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   70
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   69
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   68
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   66
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   64
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   63
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   62
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   61
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   59
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   58
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   57
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   54
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   53
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   52
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   50
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   49
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   48
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   47
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   46
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   45
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   44
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   43
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   42
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   41
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   40
      Left            =   2880
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   39
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   38
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   37
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   36
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   35
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   34
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   33
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   32
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   31
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   30
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   29
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   28
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   27
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   26
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   25
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   24
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   23
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   22
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   21
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   20
      Left            =   1440
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   19
      Left            =   720
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   18
      Left            =   720
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   17
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   16
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   15
      Left            =   720
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   14
      Left            =   720
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   13
      Left            =   720
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   11
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   10
      Left            =   720
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   9
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   8
      Left            =   0
      Stretch         =   -1  'True
      Top             =   720
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   7
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   5
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   4
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3600
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   3
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   2
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   1
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Line Line100 
      BorderColor     =   &H00000000&
      X1              =   5640
      X2              =   5520
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line99 
      BorderColor     =   &H00000000&
      X1              =   5520
      X2              =   5520
      Y1              =   5160
      Y2              =   5280
   End
   Begin VB.Line Line98 
      BorderColor     =   &H00000000&
      X1              =   5160
      X2              =   5280
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line97 
      BorderColor     =   &H00000000&
      X1              =   5280
      X2              =   5280
      Y1              =   5160
      Y2              =   5280
   End
   Begin VB.Line Line96 
      BorderColor     =   &H00000000&
      X1              =   5640
      X2              =   5520
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line95 
      BorderColor     =   &H00000000&
      X1              =   5520
      X2              =   5520
      Y1              =   5640
      Y2              =   5520
   End
   Begin VB.Line Line94 
      BorderColor     =   &H00000000&
      X1              =   5280
      X2              =   5280
      Y1              =   5640
      Y2              =   5520
   End
   Begin VB.Line Line93 
      BorderColor     =   &H00000000&
      X1              =   5160
      X2              =   5280
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line92 
      BorderColor     =   &H00000000&
      X1              =   1320
      X2              =   1200
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line91 
      BorderColor     =   &H00000000&
      X1              =   1200
      X2              =   1200
      Y1              =   5160
      Y2              =   5280
   End
   Begin VB.Line Line90 
      BorderColor     =   &H00000000&
      X1              =   1320
      X2              =   1200
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line89 
      BorderColor     =   &H00000000&
      X1              =   1200
      X2              =   1200
      Y1              =   5640
      Y2              =   5520
   End
   Begin VB.Line Line88 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   5640
      Y2              =   5520
   End
   Begin VB.Line Line87 
      BorderColor     =   &H00000000&
      X1              =   840
      X2              =   960
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line86 
      BorderColor     =   &H00000000&
      X1              =   840
      X2              =   960
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Line Line85 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   5160
      Y2              =   5280
   End
   Begin VB.Line Line84 
      BorderColor     =   &H00000000&
      X1              =   5640
      X2              =   5520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line83 
      BorderColor     =   &H00000000&
      X1              =   5520
      X2              =   5520
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line82 
      BorderColor     =   &H00000000&
      X1              =   5160
      X2              =   5280
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line81 
      BorderColor     =   &H00000000&
      X1              =   5280
      X2              =   5280
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line80 
      BorderColor     =   &H00000000&
      X1              =   5640
      X2              =   5520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line79 
      BorderColor     =   &H00000000&
      X1              =   5520
      X2              =   5520
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line78 
      BorderColor     =   &H00000000&
      X1              =   5280
      X2              =   5280
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line77 
      BorderColor     =   &H00000000&
      X1              =   5160
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line76 
      BorderColor     =   &H00000000&
      X1              =   1320
      X2              =   1200
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line75 
      BorderColor     =   &H00000000&
      X1              =   1200
      X2              =   1200
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line74 
      BorderColor     =   &H00000000&
      X1              =   1320
      X2              =   1200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line73 
      BorderColor     =   &H00000000&
      X1              =   1200
      X2              =   1200
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line72 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Line Line71 
      BorderColor     =   &H00000000&
      X1              =   840
      X2              =   960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line70 
      BorderColor     =   &H00000000&
      X1              =   840
      X2              =   960
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line69 
      BorderColor     =   &H00000000&
      X1              =   960
      X2              =   960
      Y1              =   1560
      Y2              =   1680
   End
   Begin VB.Line Line68 
      BorderColor     =   &H00000000&
      X1              =   5880
      X2              =   6000
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line67 
      BorderColor     =   &H00000000&
      X1              =   6000
      X2              =   6000
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line66 
      BorderColor     =   &H00000000&
      X1              =   5880
      X2              =   6000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line65 
      BorderColor     =   &H00000000&
      X1              =   6000
      X2              =   6000
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line64 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   4800
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line63 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   4800
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line62 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   4920
      Y2              =   4800
   End
   Begin VB.Line Line61 
      BorderColor     =   &H00000000&
      X1              =   4440
      X2              =   4560
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line60 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   4800
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line59 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   4800
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line58 
      BorderColor     =   &H00000000&
      X1              =   4440
      X2              =   4560
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line57 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line56 
      BorderColor     =   &H00000000&
      X1              =   3480
      X2              =   3360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line55 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   3360
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line54 
      BorderColor     =   &H00000000&
      X1              =   3480
      X2              =   3360
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line53 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   3360
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line52 
      BorderColor     =   &H00000000&
      X1              =   3120
      X2              =   3120
      Y1              =   4920
      Y2              =   4800
   End
   Begin VB.Line Line51 
      BorderColor     =   &H00000000&
      X1              =   3000
      X2              =   3120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line50 
      BorderColor     =   &H00000000&
      X1              =   3000
      X2              =   3120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line49 
      BorderColor     =   &H00000000&
      X1              =   3120
      X2              =   3120
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line48 
      BorderColor     =   &H00000000&
      X1              =   2040
      X2              =   1920
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line47 
      BorderColor     =   &H00000000&
      X1              =   1920
      X2              =   1920
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line46 
      BorderColor     =   &H00000000&
      X1              =   2040
      X2              =   1920
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line45 
      BorderColor     =   &H00000000&
      X1              =   1920
      X2              =   1920
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line44 
      BorderColor     =   &H00000000&
      X1              =   1680
      X2              =   1680
      Y1              =   4920
      Y2              =   4800
   End
   Begin VB.Line Line43 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   1680
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line42 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   1680
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line41 
      BorderColor     =   &H00000000&
      X1              =   1680
      X2              =   1680
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line40 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   600
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line39 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   600
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line38 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   480
      Y1              =   4800
      Y2              =   4920
   End
   Begin VB.Line Line37 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   480
      Y1              =   4440
      Y2              =   4560
   End
   Begin VB.Line Line36 
      BorderColor     =   &H00000000&
      X1              =   5880
      X2              =   6000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line35 
      BorderColor     =   &H00000000&
      X1              =   6000
      X2              =   6000
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line34 
      BorderColor     =   &H00000000&
      X1              =   5880
      X2              =   6000
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line33 
      BorderColor     =   &H00000000&
      X1              =   6000
      X2              =   6000
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line32 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   4800
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line31 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   4800
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line30 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   2760
      Y2              =   2640
   End
   Begin VB.Line Line29 
      BorderColor     =   &H00000000&
      X1              =   4440
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line28 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   4800
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line27 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   4800
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00000000&
      X1              =   4440
      X2              =   4560
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00000000&
      X1              =   4560
      X2              =   4560
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00000000&
      X1              =   3480
      X2              =   3360
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   3360
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line22 
      BorderColor     =   &H00000000&
      X1              =   3480
      X2              =   3360
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line21 
      BorderColor     =   &H00000000&
      X1              =   3360
      X2              =   3360
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line20 
      BorderColor     =   &H00000000&
      X1              =   3120
      X2              =   3120
      Y1              =   2760
      Y2              =   2640
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00000000&
      X1              =   3000
      X2              =   3120
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00000000&
      X1              =   3000
      X2              =   3120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line17 
      BorderColor     =   &H00000000&
      X1              =   3120
      X2              =   3120
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00000000&
      X1              =   2040
      X2              =   1920
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00000000&
      X1              =   1920
      X2              =   1920
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00000000&
      X1              =   2040
      X2              =   1920
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00000000&
      X1              =   1920
      X2              =   1920
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00000000&
      X1              =   1680
      X2              =   1680
      Y1              =   2760
      Y2              =   2640
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   1680
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      X1              =   1560
      X2              =   1680
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      X1              =   1680
      X2              =   1680
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   600
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   600
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   480
      Y1              =   2640
      Y2              =   2760
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   480
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   60
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   12
      Left            =   720
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   735
   End
   Begin VB.Image imgBoard 
      DragMode        =   1  'Automatic
      Height          =   735
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   735
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   15
      X1              =   6120
      X2              =   6120
      Y1              =   360
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   14
      X1              =   5400
      X2              =   5400
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   12
      X1              =   4680
      X2              =   4680
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   11
      X1              =   4680
      X2              =   4680
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   10
      X1              =   3960
      X2              =   3960
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   9
      X1              =   3960
      X2              =   3960
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   8
      X1              =   3240
      X2              =   3240
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   7
      X1              =   3240
      X2              =   3240
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   6
      X1              =   2520
      X2              =   2520
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   5
      X1              =   2520
      X2              =   2520
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   4
      X1              =   1800
      X2              =   1800
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   3
      X1              =   1800
      X2              =   1800
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   2
      X1              =   1080
      X2              =   1080
      Y1              =   3960
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   1
      X1              =   1080
      X2              =   1080
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   6840
   End
   Begin VB.Line linV 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   13
      X1              =   5400
      X2              =   5400
      Y1              =   360
      Y2              =   3240
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   9
      X1              =   360
      X2              =   6120
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   8
      X1              =   360
      X2              =   6120
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   7
      X1              =   360
      X2              =   6120
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   6
      X1              =   360
      X2              =   6120
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   5
      X1              =   360
      X2              =   6120
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   4
      X1              =   360
      X2              =   6120
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   3
      X1              =   360
      X2              =   6120
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   1
      X1              =   360
      X2              =   6120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   0
      X1              =   360
      X2              =   6120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linH 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Index           =   2
      X1              =   360
      X2              =   6120
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   3960
      X2              =   2520
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   2520
      X2              =   3960
      Y1              =   5400
      Y2              =   6840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   3960
      X2              =   2520
      Y1              =   360
      Y2              =   1800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   2520
      X2              =   3960
      Y1              =   360
      Y2              =   1800
   End
   Begin VB.Image imgRival 
      Height          =   735
      Left            =   360
      Picture         =   "frmMain.frx":8BA9
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Pieces_Initialize()
    Dim i As Integer
    For i = 0 To TotalPieces - 1
        Pieces(i).Index = i
        Select Case Left(imgPieces(i).Tag, Len(imgPieces(i).Tag) - 3)
            Case "king"
                Pieces(i).Name = King
            Case "castle"
                Pieces(i).Name = Castle
            Case "knight"
                Pieces(i).Name = Knight
            Case "rocket"
                Pieces(i).Name = Rocket
            Case "scholar"
                Pieces(i).Name = Scholar
            Case "bishop"
                Pieces(i).Name = Bishop
            Case "pawn"
                Pieces(i).Name = Pawn
        End Select
        Pieces(i).side = IIf(i > 15, TopSide, BottomSide)
    Next
    
    '--Must not be changed
    KingIndex(TopSide) = 27
    KingIndex(BottomSide) = 4
    
    Pieces_Position
End Sub

Private Sub Pieces_Position()
    Dim i As Integer, tmpIndex As Integer
    For i = 0 To TotalPieces - 1
        tmpIndex = Right(imgPieces(i).Tag, 2)
        Pieces(i).XY = CBoard(tmpIndex)
        imgPieces(i).Move imgBoard(tmpIndex).Left + 60, imgBoard(tmpIndex).Top + 60
        imgPieces(i).Visible = True
    Next
End Sub

Private Sub Agent1_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
    Dim side As Integer
    side = IIf(CharacterID = "Merlin", TopSide, BottomSide)
    AgentAvail(side) = False
End Sub

Private Sub Agent1_Show(ByVal CharacterID As String, ByVal Cause As Integer)
    Dim side As Integer
    side = IIf(CharacterID = "Merlin", TopSide, BottomSide)
    AgentAvail(side) = True
End Sub

Private Sub Form_Load()
    Pieces_Initialize
    
    Dim i As Integer
    For i = 0 To imgBoard.Count - 1
        imgBoard(i).DragMode = vbManual
        imgBoard(i).ZOrder
    Next
    For i = 0 To imgPieces.Count - 1
        imgPieces(i).DragMode = vbAutomatic
        imgPieces(i).ZOrder
    Next
    imgRival.ZOrder vbSendToBack
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DialogExp ChessMaster(BottomSide), "Thank you for playing Chinese Chess --", "Pleased"
    
    If AgentAvail(BottomSide) Then
        Do Until ChessMaster(BottomSide).Balloon.Visible = True
            DoEvents
        Loop
        Do
            DoEvents
        Loop Until ChessMaster(BottomSide).Balloon.Visible = False
    End If
End Sub

Private Sub imgBoard_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

Dim tmpP As ChessPiece, tmpNothing As ChessPiece
    tmpP = Pieces(Source.Index)
    Pieces(Source.Index).XY = CBoard(Index)
    If IsLegalMove(tmpP, tmpP.XY, CBoard(Index), tmpNothing) Then
        If Not IsCheck(Pieces(KingIndex(tmpP.side)).XY, tmpP.side) Then
            Source.Move imgBoard(Index).Left + 60, imgBoard(Index).Top + 60
            Pieces_ChangeTurn tmpP.side
            
            Pieces_Check -tmpP.side
            
        Else
            Pieces(Source.Index).XY = tmpP.XY
        End If
    Else
        Pieces(Source.Index).XY = tmpP.XY
    End If
End Sub

Private Sub imgPieces_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
'--tmpP = piece moving, tmpRestoreP = piece being killed
Dim tmpP As ChessPiece, tmpRestoreP As ChessPiece
    tmpP = Pieces(Source.Index)
    tmpRestoreP = Pieces(Index)
    '--Assume legal move
    Pieces(Source.Index).XY = Pieces(Index).XY
    '--Kill it first and check
    imgPieces(Index).Visible = False
    Pieces_Destroy Index
    If IsLegalMove(tmpP, tmpP.XY, tmpRestoreP.XY, tmpRestoreP) Then
        If Not IsCheck(Pieces(KingIndex(tmpP.side)).XY, tmpP.side) Then
            Source.Move imgBoard(CIndex(tmpRestoreP.XY)).Left + 60, imgBoard(CIndex(tmpRestoreP.XY)).Top + 60
            Pieces_ChangeTurn tmpP.side
            
            Pieces_Check -tmpP.side
            
        Else
            '--Restore order
            Pieces(Source.Index).XY = tmpP.XY
            Pieces(Index) = tmpRestoreP
            imgPieces(Index).Visible = True
        End If
    Else
        '--Restore order
        Pieces(Source.Index).XY = tmpP.XY
        Pieces(Index) = tmpRestoreP
        imgPieces(Index).Visible = True
    End If
End Sub

Private Sub Pieces_Check(side As Integer)
    If IsCheck(Pieces(KingIndex(side)).XY, side) Then
        DialogExp ChessMaster(-side), , "GetAttention"
        If IsCheckmate(Pieces(KingIndex(side)).XY, side) Then
            DialogExp ChessMaster(-side), "Checkmate!"
            If AgentAvail(-side) Then
                '--Still some problems dealing with the agent
                Do Until ChessMaster(-side).Balloon.Visible = True
                    DoEvents
                Loop
                Do
                    DoEvents
                Loop Until ChessMaster(-side).Balloon.Visible = False
            End If
            Pieces_Initialize
        Else
            DialogExp ChessMaster(-side), "Check!"
        End If
    End If
End Sub

Private Sub Pieces_ChangeTurn(side As Integer)
    Dim i As Integer, loopFrom As Integer, loopTo As Integer
    loopFrom = IIf(side = TopSide, 0, 16)
    loopTo = IIf(side = TopSide, 15, 31)
    For i = loopFrom To loopTo
        imgPieces(i).DragMode = vbAutomatic
        imgPieces(31 - i).DragMode = vbManual
    Next
End Sub

Public Sub Pieces_Destroy(Index As Integer)
    Pieces(Index).side = 0
    Pieces(Index).XY.X = 0
    Pieces(Index).XY.Y = 0
End Sub
