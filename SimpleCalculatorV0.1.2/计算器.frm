VERSION 5.00
Begin VB.Form SimpleCalculator 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "¼òµ¥¼ÆËãÆ÷"
   ClientHeight    =   3360
   ClientLeft      =   8445
   ClientTop       =   1875
   ClientWidth     =   2760
   Icon            =   "¼ÆËãÆ÷.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2760
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
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
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton point 
      Caption         =   "."
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
      TabIndex        =   19
      Top             =   2760
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
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   975
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   2280
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
      TabIndex        =   8
      Top             =   2280
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
      TabIndex        =   20
      Top             =   2280
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
      TabIndex        =   11
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
      Top             =   1800
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
      Top             =   1800
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
      Top             =   1800
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
      TabIndex        =   12
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
      Top             =   1320
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
      Top             =   1320
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
      Top             =   1320
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
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton backspace 
      Caption         =   "<"
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
      TabIndex        =   16
      Top             =   840
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
      Left            =   600
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   0
      Text            =   "txtDisplay"
      Top             =   360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   17
      Top             =   -120
      Width           =   2775
      Begin VB.TextBox txtDisplayed 
         Height          =   270
         Left            =   120
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CE 
         Caption         =   "CE"
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
         TabIndex        =   21
         Top             =   960
         Width           =   495
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Height          =   3615
         Left            =   2160
         TabIndex        =   18
         Top             =   0
         Width           =   495
         Begin VB.Label Label1 
            BackColor       =   &H8000000B&
            Caption         =   "¼òµ¥¼ÆËãÆ÷"
            BeginProperty Font 
               Name            =   "Î¢ÈíÑÅºÚ"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   120
            TabIndex        =   23
            Top             =   840
            Width           =   255
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   2760
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Menu utsidemenu 
      Caption         =   "Èý ²Ëµ¥ (&M)"
      Begin VB.Menu SimpleCalculatorBtn 
         Caption         =   "¼òµ¥¼ÆËãÆ÷"
         Enabled         =   0   'False
         Index           =   1
      End
      Begin VB.Menu BaseConversionBtn 
         Caption         =   "½øÖÆ×ª»»"
         Index           =   1
      End
      Begin VB.Menu br 
         Caption         =   "-"
      End
      Begin VB.Menu settingBtn 
         Caption         =   "ÉèÖÃ"
      End
      Begin VB.Menu br1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "ÍË³ö"
      End
   End
   Begin VB.Menu help 
      Caption         =   "°ïÖú (&H)"
      Begin VB.Menu about 
         Caption         =   "¹ØÓÚ""¼òµ¥¼ÆËãÆ÷"""
      End
   End
End
Attribute VB_Name = "SimpleCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentOperand As Double
Private calculatorOperation As String
Private isOperationPerformed As Boolean

Private Sub BaseConversionBtn_Click(Index As Integer)
    BaseConversion.Show
    Unload Me
End Sub

Private Sub Form_Load()
    Me.BorderStyle = vbFixedDialog
    currentOperand = 0
    calculatorOperation = ""
    isOperationPerformed = False
    txtDisplay.Text = "0"
    txtDisplayed.Text = " "
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

Private Sub point_Click()
    ' ÔÚÏÔÊ¾ÎÄ±¾¿òÖÐÌí¼ÓÐ¡Êýµã
    If InStr(txtDisplay.Text, ".") = 0 Then
        txtDisplay.Text = txtDisplay.Text & "."
    End If
End Sub

Private Sub CE_Click()
    txtDisplay.Text = 0
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
        txtDisplayed.Text = currentOperand
    calculatorOperation = operation
    isOperationPerformed = True
    txtDisplayed.Text = Val(txtDisplay.Text) & " " & calculatorOperation
End Sub

Private Sub btnAdd_Click()
    Operation_Click "+"
End Sub

Private Sub btnSubtract_Click()
    Operation_Click "-"
End Sub

Private Sub btnMultiply_Click()
    Operation_Click "¡Á"
End Sub

Private Sub btnDivide_Click()
    Operation_Click "¡Â"
End Sub

Private Sub Equal_Click()
    Dim result As Double

    Select Case calculatorOperation
        Case "+"
            result = currentOperand + Val(txtDisplay.Text)
        Case "-"
            result = currentOperand - Val(txtDisplay.Text)
        Case "¡Á"
            result = currentOperand * Val(txtDisplay.Text)
        Case "¡Â"
            If Val(txtDisplay.Text) <> 0 Then
                result = currentOperand / Val(txtDisplay.Text)
            Else
                MsgBox "²»ÄÜ³ýÒÔ0"
                Exit Sub
            End If
    End Select
    
    txtDisplay.Text = result
    currentOperand = result
    calculatorOperation = result
    txtDisplayed.Text = " "
    isOperationPerformed = True
End Sub

Private Sub btnEqual_Click()
    Equal_Click
End Sub

Private Sub btnClear_Click()
    txtDisplay.Text = "0"
    txtDisplayed.Text = " "
    currentOperand = 0
    calculatorOperation = ""
    isOperationPerformed = False
End Sub

Private Sub about_Click()
    Form2.Show
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub settingBtn_Click()
    setting.Show
End Sub
