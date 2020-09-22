VERSION 5.00
Begin VB.Form make5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Make 5"
   ClientHeight    =   8625
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   102
      Left            =   -240
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   103
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   101
      Left            =   11520
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   102
      Top             =   1680
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   100
      Left            =   7440
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   101
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   900
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   74
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   99
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   73
      Left            =   5568
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   98
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   72
      Left            =   4752
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   97
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   71
      Left            =   3970
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   96
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   70
      Left            =   3188
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   95
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   69
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   94
      Top             =   750
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   68
      Left            =   10464
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   93
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   67
      Left            =   9648
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   92
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   66
      Left            =   8832
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   91
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   65
      Left            =   8016
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   90
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   64
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   89
      Top             =   1560
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   63
      Left            =   6600
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   88
      Top             =   1440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   62
      Left            =   6000
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   87
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   61
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   86
      Top             =   1725
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   60
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   85
      Top             =   1740
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   59
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   84
      Top             =   2400
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   58
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   57
      Left            =   5280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   82
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   56
      Left            =   4560
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   81
      Top             =   2880
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   55
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   80
      Top             =   2370
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   54
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   79
      Top             =   3240
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   53
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   78
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   52
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   77
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   51
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   76
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   50
      Left            =   3960
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   75
      Top             =   3360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   99
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   74
      Top             =   3960
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   98
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   73
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   97
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   72
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   96
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   71
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   95
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   70
      Top             =   5280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   94
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   69
      Top             =   4800
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   93
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   68
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   92
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   67
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   91
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   66
      Top             =   4215
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   90
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   65
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   89
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   64
      Top             =   5580
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   88
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   63
      Top             =   4860
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   87
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   62
      Top             =   4860
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   86
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   61
      Top             =   4845
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   85
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   60
      Top             =   4860
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   84
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   59
      Top             =   6210
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   83
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   58
      Top             =   7440
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   82
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   57
      Top             =   6670
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   81
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   56
      Top             =   7355
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   80
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   55
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   79
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   54
      Top             =   6840
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   78
      Left            =   11280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   53
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   77
      Left            =   10440
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   52
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   76
      Left            =   9660
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   51
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   75
      Left            =   8880
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   50
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   49
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   48
      Left            =   2700
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   48
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   47
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   47
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   46
      Left            =   1140
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   46
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   45
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   45
      Top             =   4655
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   44
      Left            =   3480
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   44
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   43
      Left            =   2700
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   43
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   42
      Left            =   1920
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   42
      Top             =   4230
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   41
      Left            =   1140
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   41
      Top             =   4215
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   40
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   40
      Top             =   5320
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   39
      Left            =   3240
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   38
      Left            =   2700
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   38
      Top             =   4860
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   37
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   37
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   36
      Left            =   1140
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   36
      Top             =   4845
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   35
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   35
      Top             =   5985
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   34
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   34
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   33
      Left            =   3420
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   33
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   32
      Left            =   2640
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   32
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   31
      Left            =   1860
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   31
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   30
      Left            =   1080
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   30
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   29
      Left            =   8100
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   29
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   28
      Left            =   7320
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   28
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   27
      Left            =   6540
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   27
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   26
      Left            =   5760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   26
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   25
      Left            =   4980
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   25
      Top             =   8040
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   24
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   24
      Top             =   3990
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   23
      Left            =   1140
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   22
      Left            =   1320
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   22
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   21
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   20
      Left            =   4200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   20
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   19
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   19
      Top             =   3325
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   18
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   18
      Top             =   3960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   17
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   17
      Top             =   4800
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   16
      Left            =   2040
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   15
      Left            =   2760
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   14
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   14
      Top             =   2660
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   13
      Left            =   900
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   1605
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   12
      Left            =   2280
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   11
      Left            =   2880
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   11
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   10
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   10
      Top             =   1740
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   9
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   1995
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   8
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   1330
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   7
      Left            =   1560
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   6
      Left            =   6384
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   6
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   7200
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   5
      Top             =   30
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   3
      Left            =   60
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   665
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   2
      Left            =   842
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   2
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   1624
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox Col 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      DrawStyle       =   1  'Dash
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   2406
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Line Line1 
      Index           =   51
      X1              =   2040
      X2              =   2040
      Y1              =   -1440
      Y2              =   5880
   End
   Begin VB.Line Line1 
      Index           =   50
      X1              =   1680
      X2              =   1680
      Y1              =   -1080
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   49
      X1              =   1440
      X2              =   1440
      Y1              =   -840
      Y2              =   6480
   End
   Begin VB.Line Line1 
      Index           =   48
      X1              =   1200
      X2              =   1200
      Y1              =   -840
      Y2              =   6480
   End
   Begin VB.Line Line1 
      Index           =   47
      X1              =   1080
      X2              =   1080
      Y1              =   -720
      Y2              =   6600
   End
   Begin VB.Line Line1 
      Index           =   46
      X1              =   720
      X2              =   720
      Y1              =   -480
      Y2              =   6840
   End
   Begin VB.Line Line1 
      Index           =   45
      X1              =   360
      X2              =   360
      Y1              =   -240
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   44
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   43
      X1              =   -360
      X2              =   -360
      Y1              =   360
      Y2              =   7680
   End
   Begin VB.Line Line1 
      Index           =   42
      X1              =   -1080
      X2              =   -1080
      Y1              =   600
      Y2              =   7920
   End
   Begin VB.Line Line1 
      Index           =   41
      X1              =   13800
      X2              =   13800
      Y1              =   -5040
      Y2              =   2280
   End
   Begin VB.Line Line1 
      Index           =   40
      X1              =   13440
      X2              =   13440
      Y1              =   -4680
      Y2              =   2640
   End
   Begin VB.Line Line1 
      Index           =   39
      X1              =   13200
      X2              =   13200
      Y1              =   -4440
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   38
      X1              =   12960
      X2              =   12960
      Y1              =   -4440
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   37
      X1              =   12840
      X2              =   12840
      Y1              =   -4320
      Y2              =   3000
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   12480
      X2              =   12480
      Y1              =   -4080
      Y2              =   3240
   End
   Begin VB.Line Line1 
      Index           =   35
      X1              =   12120
      X2              =   12120
      Y1              =   -3840
      Y2              =   3480
   End
   Begin VB.Line Line1 
      Index           =   34
      X1              =   11880
      X2              =   11880
      Y1              =   -3600
      Y2              =   3720
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   11400
      X2              =   11400
      Y1              =   -3240
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   10680
      X2              =   10680
      Y1              =   -3000
      Y2              =   4320
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   9720
      X2              =   9720
      Y1              =   -3840
      Y2              =   3480
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   9360
      X2              =   9360
      Y1              =   -3480
      Y2              =   3840
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   9120
      X2              =   9120
      Y1              =   -3240
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   8880
      X2              =   8880
      Y1              =   -3240
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   8760
      X2              =   8760
      Y1              =   -3120
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   8400
      X2              =   8400
      Y1              =   -2880
      Y2              =   4440
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   8040
      X2              =   8040
      Y1              =   -2640
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   7800
      X2              =   7800
      Y1              =   -2400
      Y2              =   4920
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   7320
      X2              =   7320
      Y1              =   -2040
      Y2              =   5280
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   6600
      X2              =   6600
      Y1              =   -1800
      Y2              =   5520
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   5280
      X2              =   5280
      Y1              =   -3120
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   5520
      X2              =   5520
      Y1              =   -2280
      Y2              =   5040
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   5880
      X2              =   5880
      Y1              =   -1920
      Y2              =   5400
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   4440
      X2              =   4440
      Y1              =   -2640
      Y2              =   4680
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   4680
      X2              =   4680
      Y1              =   -3120
      Y2              =   4200
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   4920
      X2              =   4920
      Y1              =   -3240
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   5160
      X2              =   5160
      Y1              =   -3480
      Y2              =   3840
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   4320
      X2              =   4320
      Y1              =   -2040
      Y2              =   5280
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   4080
      X2              =   4080
      Y1              =   -1440
      Y2              =   5880
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   3720
      X2              =   3720
      Y1              =   -1080
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   3480
      X2              =   3480
      Y1              =   -840
      Y2              =   6480
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   3240
      X2              =   3240
      Y1              =   -840
      Y2              =   6480
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   3120
      X2              =   3120
      Y1              =   -720
      Y2              =   6600
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2760
      X2              =   2760
      Y1              =   -480
      Y2              =   6840
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2400
      X2              =   2400
      Y1              =   -240
      Y2              =   7080
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2160
      X2              =   2160
      Y1              =   0
      Y2              =   7320
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1680
      X2              =   1680
      Y1              =   360
      Y2              =   7680
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   10080
      X2              =   10080
      Y1              =   -4200
      Y2              =   3120
   End
   Begin VB.Label Comments 
      AutoSize        =   -1  'True
      Caption         =   "HOW'S THAT?!"
      BeginProperty Font 
         Name            =   "Operating instructions"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4575
      TabIndex        =   100
      Top             =   4102
      Width           =   2760
   End
   Begin VB.Menu sp 
      Caption         =   "&SPEED"
   End
End
Attribute VB_Name = "make5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Dim speeds As Variant
speeds = InputBox("Enter Speed for bulbs between 1-9999")
Timer1.Interval = speeds
For m = 0 To 50
Line1.Item(m).Visible = False
Line1.Item(m).BorderColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next m
End Sub

Private Sub sp_Click()
Dim speeds As Variant
speeds = InputBox("Enter Speed for bulbs between 1-9999")
Timer1.Interval = speeds
End Sub

Private Sub Timer1_Timer()
Dim m As Integer
For m = 0 To 50
Line1.Item(m).Visible = True
Line1.Item(m).BorderColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Next m
Dim c As Collection
Comments.ForeColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)
Dim v As Integer
For v = 0 To 99
Col.Item(v).BackColor = RGB(Rnd * 255, Rnd * 255, Rnd * 255)

Next v
End Sub
