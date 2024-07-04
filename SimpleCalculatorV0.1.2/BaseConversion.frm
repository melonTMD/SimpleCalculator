VERSION 5.00
Begin VB.Form BaseConversion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ת��"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   2760
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2760
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox targetCombo 
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox resultText 
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   5
      Text            =   "txtDisplay"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   2775
      Begin VB.CommandButton ConvertBtn 
         Caption         =   "ת��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   26.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   1935
      End
      Begin VB.ComboBox originalCombo 
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox originalText 
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Text            =   "txtDisplay"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H8000000B&
         Height          =   3615
         Left            =   2160
         TabIndex        =   1
         Top             =   0
         Width           =   495
         Begin VB.Label Label1 
            BackColor       =   &H8000000B&
            Caption         =   "����ת��"
            BeginProperty Font 
               Name            =   "΢���ź�"
               Size            =   12
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   2
            Top             =   1080
            Width           =   255
         End
      End
      Begin VB.Label Label3 
         Caption         =   "Ŀ�����:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ԭʼ����:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         X1              =   0
         X2              =   2760
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Menu Menu 
      Caption         =   "�� �˵� (&M)"
      Begin VB.Menu SimpleCalculatorBtn 
         Caption         =   "�򵥼�����"
      End
      Begin VB.Menu BaseConversionBtn 
         Caption         =   "����ת��"
         Enabled         =   0   'False
      End
      Begin VB.Menu br 
         Caption         =   "-"
      End
      Begin VB.Menu settingBtn 
         Caption         =   "����"
      End
      Begin VB.Menu br1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu help 
      Caption         =   "���� (&H)"
      Begin VB.Menu about 
         Caption         =   "����""�򵥼�����"""
      End
   End
End
Attribute VB_Name = "BaseConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    originalText = ""
    resultText = ""
        With originalCombo
        .AddItem "2"
        .AddItem "8"
        .AddItem "10"
        .AddItem "16"
        .AddItem "32"
        .AddItem "36"
        .Text = .List(2)
    End With
        With targetCombo
        .AddItem "2"
        .AddItem "8"
        .AddItem "10"
        .AddItem "16"
        .AddItem "32"
        .AddItem "36"
        .Text = .List(0)
    End With
End Sub

Private Sub SimpleCalculatorBtn_Click()
    SimpleCalculator.Show
    Unload Me
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

Private Sub ConvertBtn_Click()
    Dim originalNumber As String
    Dim originalBase As Integer
    Dim targetBase As Integer
    Dim result As String
    
    ' ��ȡԭʼ���ֺͽ���
    originalNumber = originalText.Text
    originalBase = Val(originalCombo.Text)
    targetBase = Val(targetCombo.Text)
    
    ' ת������
    result = ConvertBase(originalNumber, originalBase, targetBase)
    
    ' ��ʾ���
    resultText.Text = result
End Sub

Private Function ConvertBase(ByVal number As String, ByVal fromBase As Integer, ByVal toBase As Integer) As String
    Dim tempNumber As Double
    Dim result As String
    Dim remainder As Integer
    Dim baseChars As String
    Dim i As Integer
    
    ' ��������ַ�
    baseChars = "0123456789abcdefghijklmnopqrstuvwxyz"
    
    ' �������Ƿ���Ч
    If fromBase < 2 Or fromBase > 36 Or toBase < 2 Or toBase > 36 Then
        ConvertBase = "������Ч"
        Exit Function
    End If
    
    ' ת��Ϊʮ����
    tempNumber = 0
    For i = 1 To Len(number)
        tempNumber = tempNumber * fromBase + InStr(baseChars, Mid(number, i, 1)) - 1
    Next i
    
    ' ��ʮ����ת��ΪĿ�����
    Do While tempNumber > 0
        remainder = Int(tempNumber Mod toBase)
        result = Mid(baseChars, remainder + 1, 1) & result
        tempNumber = Int(tempNumber / toBase)
    Loop
    
    ' ����0���������
    If result = "" Then result = "0"
    
    ConvertBase = result
End Function

