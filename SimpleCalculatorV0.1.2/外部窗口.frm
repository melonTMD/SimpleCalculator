VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   4095
   ClientLeft      =   930
   ClientTop       =   2730
   ClientWidth     =   9135
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   9135
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame3 
      Caption         =   "The MIT License (MIT)"
      Height          =   3855
      Left            =   2520
      TabIndex        =   8
      Top             =   120
      Width           =   6495
      Begin VB.Label Label8 
         Caption         =   $"外部窗口.frx":0000
         Height          =   1455
         Left            =   120
         TabIndex        =   12
         Top             =   2280
         Width           =   6255
      End
      Begin VB.Label Label7 
         Caption         =   "The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software."
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   6255
      End
      Begin VB.Label Label6 
         Caption         =   $"外部窗口.frx":01D2
         Height          =   1335
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label5 
         Caption         =   "Copyright (C) 2024 melonTMD"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "使用了"
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
      Begin VB.CommandButton onGithub 
         BackColor       =   &H80000000&
         Caption         =   "在github上访问"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Windows 10"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "(C) Micosoft"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "MS Visual Basic 6.0"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Github"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   $"外部窗口.frx":0388
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "(C) Micosoft"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "关于"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      Begin VB.Label Label9 
         Caption         =   "作者脑子一抽想出来做简单计算器，所以做了这个。"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "简单计算器"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "(C) melonTMD"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub onGithub_Click()
    Dim URL As String
    URL = "https://github.com/melonTMD/SimpleCalculator"
    Shell "cmd /c start " & URL, vbHide
End Sub
