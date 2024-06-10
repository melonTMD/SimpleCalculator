VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¼òµ¥¼ÆËãÆ÷"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2775
   Icon            =   "¼ÆËãÆ÷.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2775
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton utsidemenu 
      Caption         =   "Èý Íâ²¿²Ëµ¥                         "
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   2775
   End
   Begin VB.CommandButton backspace 
      Caption         =   "ÍË¸ñ"
      BeginProperty Font 
         Name            =   "ÐÂËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton btnDivide 
      Caption         =   "¡Â"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnMultiply 
      Caption         =   "¡Á"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnSubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton btnEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   15
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnNum9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnNum8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnNum7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnNum6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton btnNum5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton btnNum4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton btnNum3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton btnNum2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton btnNum1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "ÐÂËÎÌå"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Text            =   "txtDisplay"
      Top             =   480
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   18
      Top             =   240
      Width           =   2775
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   2160
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentOperand As Double
Private calculatorOperation As String
Private isOperationPerformed As Boolean

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Me.BorderStyle = vbFixedDialog
    currentOperand = 0
    calculatorOperation = ""
    isOperationPerformed = False
    txtDisplay.Text = "0"
End Sub

Private Sub Number_Click(Index As Integer)
    If isOperationPerformed Then
        txtDisplay.Text = ""
        isOperationPerformed = False
    End If
    
    If txtDisplay.Text = "0" Then
        txtDisplay.Text = Index
    Else
        txtDisplay.Text = txtDisplay.Text & Index
    End If
End Sub

Private Sub btnNum0_Click()
    Number_Click 0
End Sub

Private Sub btnNum1_Click()
    Number_Click 1
End Sub

Private Sub btnNum2_Click()
    Number_Click 2
End Sub

Private Sub btnNum3_Click()
    Number_Click 3
End Sub

Private Sub btnNum4_Click()
    Number_Click 4
End Sub

Private Sub btnNum5_Click()
    Number_Click 5
End Sub

Private Sub btnNum6_Click()
    Number_Click 6
End Sub

Private Sub btnNum7_Click()
    Number_Click 7
End Sub

Private Sub btnNum8_Click()
    Number_Click 8
End Sub

Private Sub btnNum9_Click()
    Number_Click 9
End Sub

Private Sub backspace_Click()
    If txtDisplay.Text <> "0" Then
        If Len(txtDisplay.Text) > 1 Then
            txtDisplay.Text = Mid(txtDisplay.Text, 1, Len(txtDisplay.Text) - 1)
        Else
            txtDisplay.Text = "0"
        End If
    End If
End Sub

Private Sub Operation_Click(operation As String)
    If currentOperand <> 0 Then
        Equal_Click
    End If
    
    currentOperand = Val(txtDisplay.Text)
    calculatorOperation = operation
    isOperationPerformed = True
End Sub

Private Sub btnAdd_Click()
    Operation_Click "+"
End Sub

Private Sub btnSubtract_Click()
    Operation_Click "-"
End Sub

Private Sub btnMultiply_Click()
    Operation_Click "*"
End Sub

Private Sub btnDivide_Click()
    Operation_Click "/"
End Sub

Private Sub Equal_Click()
    Dim result As Double
    Select Case calculatorOperation
        Case "+"
            result = currentOperand + Val(txtDisplay.Text)
        Case "-"
            result = currentOperand - Val(txtDisplay.Text)
        Case "*"
            result = currentOperand * Val(txtDisplay.Text)
        Case "/"
            If Val(txtDisplay.Text) <> 0 Then
                result = currentOperand / Val(txtDisplay.Text)
            Else
                MsgBox "²»ÄÜ³ýÒÔ0"
                Exit Sub
            End If
    End Select
    
    txtDisplay.Text = result
    currentOperand = result
    calculatorOperation = ""
    isOperationPerformed = True
End Sub

Private Sub btnEqual_Click()
    Equal_Click
End Sub

Private Sub btnClear_Click()
    txtDisplay.Text = "0"
    currentOperand = 0
    calculatorOperation = ""
    isOperationPerformed = False
End Sub

Private Sub utsidemenu_Click()
    Form2.Show
End Sub
