VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   6135
   ClientLeft      =   7980
   ClientTop       =   2700
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command12 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   1920
      TabIndex        =   22
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   1080
      TabIndex        =   21
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1920
      TabIndex        =   19
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1080
      TabIndex        =   18
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   17
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   16
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "BS"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Off"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Minimal Calc"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command3 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "On"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   5520
         Width           =   3135
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim operatorFlag As Boolean
Dim temp As Double
Dim operatorSign As Integer


Private Sub Command1_Click()
Dim i As Integer
    i = 0
    Text1.Text = ""
    Text1.BackColor = &HC0FFC0
    While i <= 9
        Command12(i).Enabled = True
        i = i + 1
    Wend
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True
    Command7.Enabled = True
    Command8.Enabled = True
    Command9.Enabled = True
    Command10.Enabled = True
    Command11.Enabled = True
End Sub

Private Sub Command10_Click()
    Select Case operatorSign
        Case 1
            Text1.Text = temp + Val(Text1.Text)
        Case 2
            Text1.Text = temp - Val(Text1.Text)
        Case 3
            Text1.Text = temp * Val(Text1.Text)
        Case 4
            Text1.Text = temp / Val(Text1.Text)
        Case 5
            Text1.Text = temp Mod Val(Text1.Text)
    End Select
End Sub

Private Sub Command11_Click()
    Text1.Text = Text1.Text & "."
End Sub

Private Sub Command12_Click(Index As Integer)
    If operatorFlag = False Then
        Text1.Text = Text1.Text & Index
    Else
        temp = Text1.Text
        Text1.Text = Index
        operatorFlag = False
    End If
End Sub

Private Sub Command2_Click()
    Dim i As Integer
    i = 0
    Text1.Text = ""
    Text1.BackColor = &H80000011
    While i <= 9
        Command12(i).Enabled = False
        i = i + 1
    Wend
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command11.Enabled = False
    operatorFlag = False
End Sub

Private Sub Command3_Click()
    Text1.Text = ""
End Sub

Private Sub Command4_Click()
    Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
End Sub

Private Sub Command5_Click()
    operatorFlag = True
    operatorSign = 5
End Sub

Private Sub Command6_Click()
    operatorFlag = True
    operatorSign = 4
End Sub

Private Sub Command7_Click()
    operatorFlag = True
    operatorSign = 3
End Sub

Private Sub Command8_Click()
    operatorFlag = True
    operatorSign = 2
End Sub

Private Sub Command9_Click()
    operatorFlag = True
    operatorSign = 1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    i = 0
    Text1.Text = ""
    Text1.BackColor = &H80000011
    While i <= 9
        Command12(i).Enabled = False
        i = i + 1
    Wend
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    Command7.Enabled = False
    Command8.Enabled = False
    Command9.Enabled = False
    Command10.Enabled = False
    Command11.Enabled = False
    operatorFlag = False
End Sub
